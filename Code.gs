// ═══════════════════════════════════════════════════════════════
// ENTERPRISE BUSINESS CONTROL TOWER — Apps Script Backend
// Reads the Sellers + Buyers sheets server-side and serves the
// audience-aware dashboard data. Everything is derived from the sheet
// columns (business_vertical, business_category, onboarding_status,
// gstin_status, transaction_activation_status, dates); nothing hard-coded.
//
// Because the sheets are read with SpreadsheetApp (the deploying account's
// access), they can stay PRIVATE — no public sharing and no CORS.
// ═══════════════════════════════════════════════════════════════

var CONFIG = {
  CACHE_TTL: 300,  // seconds — 5-minute cache; matches Metabase sync frequency

  // Metabase card IDs — keyed by audience.
  // readData() prefers the _mb_* sheets synced by the Metabase.gs trigger; when
  // those are absent it calls fetchMBCard_() (Metabase.gs) directly, so the
  // dashboard is live even before the first sync. Requires MB credentials in
  // Script Properties (setupMetabaseCredentials in Metabase.gs).
  MB_CARDS: {
    seller: 5712,  // Seller Info
    buyer:  5711,  // Buyer from Inspection
  },

  // Hidden sheet names written by Metabase.gs into THIS spreadsheet.
  // Serve as a warm fallback when the Metabase API is temporarily unreachable.
  MB_LOCAL_SHEETS: {
    seller: '_mb_sellers',
    buyer:  '_mb_buyers',
  },

  // Last-resort fallback sheet IDs — used only when Metabase is unavailable
  // and no local synced sheets exist yet.  Dashboard always loads with real data.
  SELLER_SHEET_ID: '1qaG_GMvUrC7LJbKBma8-S3x2N6Ua4vDkUC5ZRx3znho',
  BUYER_SHEET_ID:  '1G9Ocq8PXovCx5eBE3dOfE8CUcmfgwcl6Te37p79u0wI',

  // TAT source — dedicated tab holding Metabase card 5292 ("Onboarding Detail"),
  // synced by Metabase.gs (syncDetail) every 5 min. Every TAT metric is derived
  // from this tab's "In Review" column; see getTATLookup_().
  TAT_DETAIL_SHEET: '_mb_detail',
};

// Candidate column names in the 5292 detail tab (headers are normalized to
// lower_snake_case by readSheetObj_). First populated match wins per record.
var TAT_COLS = {
  inReview:  ['in_review_date','in_review_at','in_review_on','in_review','inreview_date',
              'inreview','review_date','review_at','under_review_date','submitted_date','submitted_at'],
  // Fallback start when In Review is blank.
  created:   ['onboarding_created_date','created_date','created_at','created_on',
              'onboarding_created_at','registration_date','signup_date'],
  completed: ['onboarding_completed_date','completed_date','completion_date','onboarded_date',
              'completed_at','onboarded_at','onboarding_updated_date','completed_on'],
  status:    ['onboarding_status','status','onboarding_state','state'],
  id:        ['seller_id','buyer_id','vendor_id','onboarding_id','id','entity_id','user_id','account_id'],
  name:      ['seller_name','buyer_name','vendor_name','business_name','name','entity_name'],
};

// Per-audience column mapping (field names differ slightly between Metabase cards and the fallback sheets).
var AUDIENCE_CFG = {
  seller: { audience: 'seller', sheetId: CONFIG.SELLER_SHEET_ID, label: 'Sellers',
            idCol: 'seller_id', nameCol: 'seller_name', vendorCol: 'vendor_type', onbCol: 'onboarded_date' },
  // Buyer card has no onboarded_date; TAT uses onboarding_updated_date (last status change).
  buyer:  { audience: 'buyer',  sheetId: CONFIG.BUYER_SHEET_ID,  label: 'Buyers',
            idCol: 'buyer_id',  nameCol: 'buyer_name',  vendorCol: 'customer_type', onbCol: 'onboarding_updated_date' },
};

// TAT beyond this many days is treated as data noise (bulk re-import timestamps)
// and excluded from the Avg TAT so a few outliers don't skew it.
var TAT_MAX_DAYS = 365;

// The eight business verticals, in display order. Rows are mapped into these by
// mapToVertical() + date-based migration in normalizeRows.
var VERTICALS = [
  { key: 'OMP',           name: 'Open Marketplace', code: 'OMP', sub: '' },
  { key: 'EPR',           name: 'EPR',              code: 'EPR', sub: '' },
  { key: 'Marketplace',   name: 'Marketplace',      code: 'MKT', sub: '' },
  { key: 'InfraBusiness', name: 'Infra Business',   code: 'INF', sub: '' },
  { key: 'AFR',           name: 'AFR',              code: 'AFR', sub: '' },
  { key: 'Recommerce',    name: 'Re-Commerce',      code: 'REC', sub: '' },
  { key: 'DRS',           name: 'DRS',              code: 'DRS', sub: '' },
  { key: 'Others',        name: 'Others',           code: 'OTH', sub: '' },
];

// Categories (from the Marketplace business_vertical) that roll into Infra Business.
var INFRA_CATS = { 'metal': 1, 'plastic': 1, 'institutional business': 1, 'reverse': 1, 'rewerse': 1 };
function isEwaste(cat) { return cat === 'e-waste' || cat === 'ewaste' || cat === 'e waste'; }

// Marketplace infra rows (Metal/Plastic/Institutional/Reverse) split at FY 26-27:
//   created < April 1 2026  → 'Marketplace' card (historical)
//   created >= April 1 2026 → 'InfraBusiness' card (current)
var MKT_SPLIT_DATE = new Date(2026, 3, 1); // April 1, 2026

// Map a sheet row to one of the eight verticals (date-unaware for Marketplace infra —
// normalizeRows applies the FY 26-27 split after this call).
//  · EPR              ← business_vertical 'EPR' (all of it)
//  · Open Marketplace ← business_vertical 'Open Marketplace', Metal or Plastic → OMP (unchanged)
//  · Marketplace vertical, by business_category:
//        AFR                                          → AFR
//        GOA DRS                                      → DRS
//        Re-Commerce, E-Waste                         → Recommerce
//        Metal, Plastic, Institutional Business, Reverse → InfraBusiness (then split by date)
//        anything else                                → Others
//  · everything else (Sustainability Services, blank, …) → Others
// audience = 'seller'|'buyer'. E-Waste under Marketplace maps to Recommerce for both.
function mapToVertical(businessVertical, category, audience) {
  var bv  = String(businessVertical || '').trim().toLowerCase();
  var cat = String(category || '').trim().toLowerCase();
  if (bv === 'epr') return 'EPR';
  if (bv === 'open marketplace' || bv === 'openmarketplace') {
    return 'OMP';  // All Open Marketplace entries belong to OMP regardless of category
  }
  if (bv === 'marketplace') {
    if (cat === 'afr') return 'AFR';
    if (cat === 'goa drs' || cat === 'drs') return 'DRS';
    if (cat === 're-commerce' || cat === 'recommerce' || cat === 're commerce') return 'Recommerce';
    if (isEwaste(cat)) return 'Recommerce';
    if (INFRA_CATS[cat]) return 'InfraBusiness';
    return 'Others';
  }
  return 'Others';
}

// ─────────────────────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enterprise Business Control Tower')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData(filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var audience = (f.audience === 'buyer') ? 'buyer' : 'seller';
    var cfg = AUDIENCE_CFG[audience];

    var periodKey = JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cacheKey  = 'dash_v30_' + audience + '_' + periodKey;
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    var raw  = readData(audience);
    var rows = normalizeRows(raw, cfg);

    // buildDashboard must run first — it mutates OMP rows in-place to set isOldVendor.
    // Caching vertical rows before this call would store rows without isOldVendor, so
    // the Existing vs New drill-down would always be wrong.
    var dash = buildDashboard(audience, cfg, rows, f);
    dash.success = true;

    // Annotate data source so the frontend can show "Live (Metabase)" vs "Sheets"
    dash.dataSource = raw.source || 'sheets';
    if (raw.source === 'metabase_live') {
      dash.mbSyncedAt = new Date().toISOString();
    } else if (raw.source === 'metabase_sync') {
      try { dash.mbSyncedAt = getMBSyncTime(CONFIG.MB_LOCAL_SHEETS[audience]); } catch (e) {}
    }

    // Pre-populate per-vertical row caches NOW (after buildDashboard has set isOldVendor
    // on OMP rows) so getVerticalRows() always gets a correct cache hit.
    var filtered = rows.filter(function(r) { return applyDateFilter(r, f); });
    var batchCache = {};
    VERTICALS.forEach(function(vc) {
      var vrows = filtered.filter(function(r) { return r.vertical === vc.key; })
        .sort(function(a, b) {
          return (b.createdDate ? b.createdDate.getTime() : 0) - (a.createdDate ? a.createdDate.getTime() : 0);
        });
      var v = JSON.stringify({ success: true, vertKey: vc.key, rows: vrows.map(vertRow) });
      if (v.length <= 100000) batchCache['vrows_v14_' + audience + '_' + vc.key + '_' + periodKey] = v;
    });

    var out = JSON.stringify(dash);
    if (out.length <= 100000) batchCache[cacheKey] = out;
    // Single RPC to write all vertical + dashboard caches at once
    try { cache.putAll(batchCache, CONFIG.CACHE_TTL); } catch (e) {}
    // Fallback: ensure the dashboard key is cached even if batch failed or was oversized
    if (!batchCache[cacheKey]) try { cache.put(cacheKey, out, CONFIG.CACHE_TTL); } catch (e) {}
    return out;
  } catch (err) {
    return JSON.stringify({ success: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// Returns seller + buyer dashboard data in one call so the frontend can
// render both pipelines side-by-side without two round trips.
function getCombinedDashboard(filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var periodKey = JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cacheKey  = 'dash_v30_cmb_' + periodKey;
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    var sRaw  = readData('seller');
    var bRaw  = readData('buyer');
    var sCfg  = AUDIENCE_CFG.seller;
    var bCfg  = AUDIENCE_CFG.buyer;
    var sDash = buildDashboard('seller', sCfg, normalizeRows(sRaw, sCfg), f);
    var bDash = buildDashboard('buyer',  bCfg, normalizeRows(bRaw, bCfg), f);

    var out = JSON.stringify({
      success:        true,
      lastUpdated:    new Date().toISOString(),
      dataSource:     sRaw.source || 'sheets',
      verticalConfig: sDash.verticalConfig,
      seller: { verticals: sDash.verticals, fiscalYears: sDash.fiscalYears },
      buyer:  { verticals: bDash.verticals, fiscalYears: bDash.fiscalYears },
    });

    try { cache.put(cacheKey, out, CONFIG.CACHE_TTL); } catch (e) {}
    return out;
  } catch (err) {
    return JSON.stringify({ success: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// On-demand row fetch for a single vertical's detail panel.
// Kept separate so the main getCombinedDashboard response stays small
// enough for Script Cache (100KB limit per value).
function getVerticalRows(vertKey, filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var audience = (f.audience === 'buyer') ? 'buyer' : 'seller';
    var cfg = AUDIENCE_CFG[audience];

    var cacheKey = 'vrows_v14_' + audience + '_' + vertKey + '_'
      + JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    var raw  = readData(audience);
    var all  = normalizeRows(raw, cfg);

    // For OMP, replicate the isOldVendor logic from buildDashboard so the drill-down
    // shows correct New vs Existing flags even when called independently.
    if (vertKey === 'OMP') {
      var existingNames = {};
      all.forEach(function(r) {
        if (r.vertical !== 'OMP' && r.status === 'COMPLETED' && r.name) {
          existingNames[r.name.trim().toLowerCase()] = true;
        }
      });
      all.forEach(function(r) {
        if (r.vertical === 'OMP') {
          r.isOldVendor = !!(r.name && existingNames[r.name.trim().toLowerCase()]);
        }
      });
    }

    var vrows = all.filter(function(r) {
      return r.vertical === vertKey && applyDateFilter(r, f);
    }).sort(function(a, b) {
      return (b.createdDate ? b.createdDate.getTime() : 0) - (a.createdDate ? a.createdDate.getTime() : 0);
    });

    var out = JSON.stringify({ success: true, vertKey: vertKey, rows: vrows.map(vertRow) });
    try { cache.put(cacheKey, out, CONFIG.CACHE_TTL); } catch (e) {}
    return out;
  } catch (err) {
    return JSON.stringify({ success: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// Returns transacted OMP vendors (both seller + buyer) for the drill-down table.
// Uses the vertical-rows cache populated by getDashboardData so no extra reads occur.
function getTransactedVendors(filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var periodKey = JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cache = CacheService.getScriptCache();

    function fetchOmpTxn(aud) {
      var cacheKey = 'vrows_v14_' + aud + '_OMP_' + periodKey;
      var hit = cache.get(cacheKey);
      if (hit) {
        try {
          var d = JSON.parse(hit);
          if (d && d.rows) return d.rows.filter(function(r) { return r.hasTransacted; });
        } catch (e) {}
      }
      // Cache miss: read from source
      var cfg = AUDIENCE_CFG[aud];
      var raw = readData(aud);
      var all = normalizeRows(raw, cfg);
      return all.filter(function(r) {
        return r.vertical === 'OMP' && r.hasTransacted && applyDateFilter(r, f);
      }).map(vertRow);
    }

    return JSON.stringify({
      success: true,
      sellers: fetchOmpTxn('seller'),
      buyers:  fetchOmpTxn('buyer'),
    });
  } catch (err) {
    return JSON.stringify({ success: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// ─────────────────────────────────────────────────────────────
// DATA READ + NORMALIZATION
// ─────────────────────────────────────────────────────────────
// Priority (freshest live data first):
//   1. _mb_* hidden sheets in this spreadsheet — pure in-process read, ~100 ms.
//      Written every 5 min by the Metabase.gs trigger; same freshness as the cache.
//   2. Metabase API direct fetch — ~3-8 s HTTP round-trip; keeps the dashboard
//      connected to live backend data before the first trigger sync has run.
//   3. External Google Sheet by ID — last-resort static fallback so the
//      dashboard still renders when Metabase is unreachable.
function readData(audience) {
  var localName  = CONFIG.MB_LOCAL_SHEETS[audience];
  var fallbackId = AUDIENCE_CFG[audience] && AUDIENCE_CFG[audience].sheetId;
  var cardId     = CONFIG.MB_CARDS[audience];

  // 1. Local synced sheet — no HTTP, fastest path
  if (localName) {
    try {
      var local = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(localName);
      if (local && local.getLastRow() > 1) {
        var d = readSheetObj_(local);
        d.source = 'metabase_sync';
        return d;
      }
    } catch (e) { /* fall through */ }
  }

  // 2. Metabase direct — live backend data when the local sync hasn't run yet
  if (cardId) {
    try {
      var mbData = fetchMBCard_(cardId);
      if (mbData && mbData.headers.length && mbData.rows.length) {
        mbData.source = 'metabase_live';
        return mbData;
      }
    } catch (e) { Logger.log('Metabase direct fetch failed: ' + e.message); }
  }

  // 3. External Google Sheet — static fallback, only when Metabase is unavailable
  if (fallbackId) {
    try {
      var fb = readSheetObj_(SpreadsheetApp.openById(fallbackId).getSheets()[0]);
      fb.source = 'sheets';
      return fb;
    } catch (e) { Logger.log('Fallback sheet read failed: ' + e.message); }
  }

  return { headers: [], rows: [], source: 'empty' };
}

function readSheetObj_(sheet) {
  var vals = sheet.getDataRange().getValues();
  if (!vals.length) return { headers: [], rows: [] };
  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  var rows = vals.slice(1).filter(function(row) {
    return row.some(function(c) { return c !== '' && c !== null && c !== undefined; });
  });
  return { headers: headers, rows: rows };
}

// ─────────────────────────────────────────────────────────────
// TAT LOOKUP — sourced exclusively from Metabase card 5292
// ("Onboarding Detail"), synced into the _mb_detail tab every 5 min.
//
// TAT per record = whole days from the "In Review" date (or the Created
// date when In Review is blank) up to the completion date (when the record
// is onboarded) or, for records still in flight, up to the current date.
// Rebuilt on every request, so the running TAT of open records advances
// automatically and refreshes whenever the 5292 sync updates the tab.
// ─────────────────────────────────────────────────────────────
var _tatLookupCache = null;   // per-execution cache; each request re-evaluates the module

// First value among candidate columns that both EXISTS and is non-empty, so an
// empty leading column (e.g. a blank generic `id`) never shadows a populated one.
function firstVal_(row, idx, keys) {
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (idx[k] !== undefined) {
      var v = row[idx[k]];
      if (v !== '' && v !== null && v !== undefined) return v;
    }
  }
  return '';
}

function getTATLookup_() {
  if (_tatLookupCache) return _tatLookupCache;
  var lk = { byId: {}, byName: {}, loaded: false, count: 0 };
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TAT_DETAIL_SHEET);
    if (sheet && sheet.getLastRow() > 1) {
      var d   = readSheetObj_(sheet);
      var idx = buildIndex(d.headers);
      var now = new Date();
      d.rows.forEach(function(row) {
        // TAT clock starts at the In Review date; when that is blank, fall back
        // to the Created date so records that skipped/haven't logged review still
        // get a TAT. No start date at all → record has no TAT yet.
        var start = parseDate(firstVal_(row, idx, TAT_COLS.inReview))
                 || parseDate(firstVal_(row, idx, TAT_COLS.created));
        if (!start) return;
        var status = normStatus(firstVal_(row, idx, TAT_COLS.status));
        var comp   = parseDate(firstVal_(row, idx, TAT_COLS.completed));
        var isDone = (status === 'COMPLETED') || (status === 'UNKNOWN' && !!comp);
        // End the clock at the completion date when onboarded; otherwise "now".
        var end = (isDone && comp) ? comp : now;
        var tat = dateDiffDays(start, end);
        if (tat === null || tat < 0) return;
        var id   = String(firstVal_(row, idx, TAT_COLS.id) || '').replace(/,/g, '').trim();
        var name = String(firstVal_(row, idx, TAT_COLS.name) || '').trim().toLowerCase();
        // On duplicate keys, a settled (completed) TAT wins over a running one.
        if (id) {
          var ex = lk.byId[id];
          if (!ex || (isDone && !ex.done)) lk.byId[id] = { tat: tat, done: isDone };
        }
        if (name) {
          var exn = lk.byName[name];
          if (!exn || (isDone && !exn.done)) lk.byName[name] = { tat: tat, done: isDone };
        }
        lk.count++;
      });
      lk.loaded = lk.count > 0;
    }
  } catch (e) { Logger.log('TAT lookup build failed: ' + e.message); }
  _tatLookupCache = lk;
  return lk;
}

// Resolve a record's TAT from the 5292 lookup: match by id first, then name.
//   undefined → detail tab unavailable (caller falls back to legacy TAT)
//   null      → tab loaded but no match for this record (excluded from TAT)
//   number    → days in review
function lookupTAT_(id, name) {
  var lk = getTATLookup_();
  if (!lk.loaded) return undefined;
  if (id)   { var e  = lk.byId[String(id).replace(/,/g, '').trim()];  if (e)  return e.tat; }
  if (name) { var en = lk.byName[String(name).trim().toLowerCase()];  if (en) return en.tat; }
  return null;
}

// Map each sheet row to the common record shape the dashboard renders from.
function normalizeRows(raw, cfg) {
  var idx = buildIndex(raw.headers);
  var normalized = raw.rows.map(function(row) {
    var status   = normStatus(gv(row, idx, 'onboarding_status'));
    var category = normCategory(gv(row, idx, 'business_category'));
    var bizVert  = gv(row, idx, 'business_vertical');
    var created  = parseDate(gv(row, idx, 'onboarding_created_date'));
    var onboarded = cfg.onbCol ? parseDate(gv(row, idx, cfg.onbCol)) : null;
    var recId    = String(gv(row, idx, cfg.idCol)   || '').replace(/,/g, '').trim();
    var recName  = String(gv(row, idx, cfg.nameCol) || '').trim();
    // TAT is sourced from Metabase card 5292 (the _mb_detail tab): days from the
    // record's "In Review" date to its completion date (or "now" while in flight),
    // matched by id then name. When that tab is unavailable, fall back to the legacy
    // created→onboarded span so Avg TAT still renders (graceful degradation).
    var lkTat = lookupTAT_(recId, recName);
    var tat = (lkTat !== undefined)
      ? lkTat
      : ((status === 'COMPLETED' && created && onboarded) ? dateDiffDays(created, onboarded) : null);

    var gstin     = String(gv(row, idx, 'gst_number') || gv(row, idx, 'gstin') || '').trim();
    var gstStatus = String(gv(row, idx, 'gstin_status') || gv(row, idx, 'gst_status') || '').toUpperCase();
    var txnRaw    = gv(row, idx, 'transaction_activation_status')
                 || gv(row, idx, 'transacted')
                 || gv(row, idx, 'is_transacted')
                 || gv(row, idx, 'transaction_status')
                 || gv(row, idx, 'txn_status')
                 || '';
    var txn = String(txnRaw).toUpperCase().replace(/\s+/g, ' ').trim();

    // Transaction value — exact matches first, then fuzzy scan for any column whose
    // name contains 'gmv' or '(transaction|txn|order)_(value|amount)'.
    var txnValRaw = (function() {
      var exact = ['transaction_value', 'txn_value', 'total_transaction_value',
                   'gmv', 'transaction_amount', 'first_transaction_value', 'order_value'];
      for (var _i = 0; _i < exact.length; _i++) {
        if (idx[exact[_i]] !== undefined) return row[idx[exact[_i]]];
      }
      var hkeys = Object.keys(idx);
      for (var _j = 0; _j < hkeys.length; _j++) {
        if (hkeys[_j].indexOf('gmv') !== -1) return row[idx[hkeys[_j]]];
      }
      for (var _k = 0; _k < hkeys.length; _k++) {
        var _h = hkeys[_k];
        if ((_h.indexOf('transaction') !== -1 || _h.indexOf('txn') !== -1 || _h.indexOf('order') !== -1)
            && (_h.indexOf('value') !== -1 || _h.indexOf('amount') !== -1)) return row[idx[_h]];
      }
      return '';
    }());
    var _txnN = (txnValRaw !== '' && txnValRaw !== null && txnValRaw !== undefined)
      ? parseFloat(String(txnValRaw).replace(/[₹$€£¥,\s]/g, ''))
      : NaN;
    var txnVal = isNaN(_txnN) ? null : _txnN; // preserve 0 — parseFloat('0')||null would lose it
    // GMV present & positive ⇒ a transaction has happened, regardless of what the
    // status column says (or if it's missing). This is the authoritative signal.
    var gmvTransacted = txnVal !== null && txnVal > 0;
    var txnDate = parseDate(
      gv(row, idx, 'first_transaction_date') ||
      gv(row, idx, 'transaction_date')       ||
      gv(row, idx, 'first_txn_date')         ||
      gv(row, idx, 'txn_date')               ||
      gv(row, idx, 'activation_date')        || ''
    );

    var vertical = mapToVertical(bizVert, category, cfg.audience);
    // Marketplace infra rows: before April 1 2026 → 'Marketplace' card (historical);
    // from April 1 2026 → 'InfraBusiness' (default from mapToVertical).
    if (vertical === 'InfraBusiness' && String(bizVert || '').trim().toLowerCase() === 'marketplace'
        && created && created < MKT_SPLIT_DATE) {
      vertical = 'Marketplace';
    }
    return {
      id:            recId,
      name:          recName,
      state:         String(gv(row, idx, 'state') || '').trim(),
      vendorType:    String(gv(row, idx, cfg.vendorCol) || '').trim(),
      category:      category,
      vertical:      vertical,
      bizVertical:   String(bizVert || '').trim(),
      gstin:         gstin,
      status:        status,
      createdDate:   created,
      onboardedDate: onboarded,
      // gstStatus is authoritative when present.  Positive signals: AVAILABLE or ACTIVE
      // (catches "GSTIN AVAILABLE", "ACTIVE", "GST ACTIVE").  Negative qualifiers that
      // override a positive: NOT, MISSING, INACTIVE (catches "NOT AVAILABLE",
      // "INACTIVE", "GST INACTIVE", "MISSING GSTIN").  Falls back to GSTIN format
      // validation when the column is absent (e.g. raw 15-char GSTIN in gst_number).
      hasGST: gstStatus
        ? ((gstStatus.indexOf('AVAILABLE') !== -1 || gstStatus.indexOf('ACTIVE') !== -1 ||
            gstStatus === 'YES' || gstStatus === 'Y' || gstStatus === 'REGISTERED' || gstStatus === 'VALID')
           && gstStatus.indexOf('NOT') === -1 && gstStatus.indexOf('MISSING') === -1 && gstStatus.indexOf('INACTIVE') === -1)
        : isValidGSTIN(gstin),
      // Vertical tracks transactions if any signal exists: status col, GMV, or a transaction date.
      hasTxn:        txn !== '' || txnVal !== null || txnDate !== null,
      // Transacted if the status says so, OR positive GMV exists, OR a transaction date is present.
      hasTransacted: (({ 'TRANSACTED': 1, 'YES': 1, 'Y': 1, 'TRUE': 1, '1': 1, 'DONE': 1,
                         'ACTIVE': 1, 'TRANSACTED YES': 1, 'COMPLETED': 1, 'SUCCESS': 1,
                         'ENABLED': 1, 'ACTIVATED': 1, 'CONFIRMED': 1 })[txn] === 1)
                  || gmvTransacted
                  || txnDate !== null,
      txnValue:      txnVal,
      txnDate:       txnDate,
      onbTAT:        tat,
    };
  }).filter(function(r) { return r.id || r.name; });
  // Count pre-dedup rows per (id+vertical) as per-vendor transaction count.
  // Metabase emits one row per transaction via joins; this count captures that.
  var txnCounts = {};
  normalized.forEach(function(r) {
    if (!r.id) return;
    var key = r.id + '\x00' + r.vertical;
    txnCounts[key] = (txnCounts[key] || 0) + 1;
  });
  // Deduplicate by ID — Metabase can emit the same entity multiple times when joins
  // produce multiple transaction rows. Sort so hasTransacted=true rows come first so
  // the dedup keeps the transacted version; break ties by txnValue descending.
  normalized.sort(function(a, b) {
    return (b.hasTransacted ? 1 : 0) - (a.hasTransacted ? 1 : 0)
        || (b.txnValue || 0) - (a.txnValue || 0);
  });
  // Dedup key = id + vertical so the same entity can appear in multiple verticals
  // (e.g. a seller in both OMP and Marketplace keeps both rows for Existing vs New).
  // Intra-vertical duplicates (same id, same vertical — Metabase join artifacts) are still removed.
  var seen = {};
  return normalized.filter(function(r) {
    if (!r.id) return true;
    var key = r.id + '\x00' + r.vertical;
    if (seen[key]) return false;
    seen[key] = true;
    r.txnCount = txnCounts[key] || 1;
    return true;
  });
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD ASSEMBLY
// ─────────────────────────────────────────────────────────────
function buildDashboard(audience, cfg, allRows, f) {
  var rows = allRows.filter(function(r) { return applyDateFilter(r, f); });

  var byVert = {};
  rows.forEach(function(r) { (byVert[r.vertical] = byVert[r.vertical] || []).push(r); });

  // Existing vs New (OMP only): a vendor is "existing" if their name appears with
  // COMPLETED status in ANY non-OMP vertical. Matched by name (case-insensitive
  // trim) because the same entity may carry different IDs across verticals.
  // Scans allRows (unfiltered) so historical completions from any period count.
  var existingNames = {};
  allRows.forEach(function(r) {
    if (r.vertical !== 'OMP' && r.status === 'COMPLETED' && r.name) {
      existingNames[r.name.trim().toLowerCase()] = true;
    }
  });
  (byVert['OMP'] || []).forEach(function(r) {
    r.isOldVendor = !!(r.name && existingNames[r.name.trim().toLowerCase()]);
  });

  // Always emit all seven verticals, in fixed order (empty ones render "No records").
  var verticalConfig = VERTICALS.map(function(vc) {
    return { key: vc.key, name: vc.name, sub: vc.sub, code: vc.code };
  });
  var verticals = {};
  VERTICALS.forEach(function(vc) {
    verticals[vc.key] = vStats(byVert[vc.key] || []);
  });

  return {
    success: true,
    lastUpdated: new Date().toISOString(),
    audience: audience,
    verticalConfig: verticalConfig,
    verticals: verticals,
    fiscalYears: buildFiscalYears(allRows),
  };
}

// ─── DIAGNOSTIC — run from Apps Script editor → View logs ───────────────────
function debugExistingVsNew() {
  Logger.log('=== Existing vs New Debug ===');
  ['seller', 'buyer'].forEach(function(aud) {
    Logger.log('\n--- ' + aud.toUpperCase() + ' ---');
    var cfg  = AUDIENCE_CFG[aud];
    var raw  = readData(aud);
    Logger.log('Data source: ' + raw.source + '  Raw rows: ' + raw.rows.length);
    var rows = normalizeRows(raw, cfg);
    Logger.log('Normalized rows (post-dedup): ' + rows.length);

    // Raw business_vertical values (before mapping) — confirms what the card returns
    var bvCounts = {};
    rows.forEach(function(r) {
      var bv = String(r.bizVertical || '(blank)').trim() || '(blank)';
      bvCounts[bv] = (bvCounts[bv] || 0) + 1;
    });
    Logger.log('Rows by raw business_vertical:');
    Object.keys(bvCounts).sort().forEach(function(k) { Logger.log('  "' + k + '" → ' + bvCounts[k]); });

    // Rows by mapped vertical
    var vCounts = {};
    rows.forEach(function(r) { vCounts[r.vertical] = (vCounts[r.vertical] || 0) + 1; });
    Logger.log('Rows by mapped vertical:');
    Object.keys(vCounts).sort().forEach(function(k) { Logger.log('  ' + k + ' → ' + vCounts[k]); });

    // Sample OMP categories (to verify category values)
    var ompRows = rows.filter(function(r) { return r.vertical === 'OMP'; });
    var ompCats = {};
    ompRows.forEach(function(r) { ompCats[r.category] = (ompCats[r.category] || 0) + 1; });
    Logger.log('OMP category breakdown:');
    Object.keys(ompCats).sort().forEach(function(k) { Logger.log('  "' + k + '" → ' + ompCats[k]); });

    // Names COMPLETED in any non-OMP vertical
    var existingNames = {};
    rows.forEach(function(r) {
      if (r.vertical !== 'OMP' && r.status === 'COMPLETED' && r.name) {
        existingNames[r.name.trim().toLowerCase()] = true;
      }
    });
    Logger.log('Non-OMP COMPLETED unique names: ' + Object.keys(existingNames).length);
    if (Object.keys(existingNames).length > 0) {
      Logger.log('  Sample: ' + Object.keys(existingNames).slice(0, 5).join(' | '));
    } else {
      Logger.log('  ⚠ ZERO non-OMP COMPLETED names — card may only contain OMP data.');
      Logger.log('  ⚠ If so, use card 5292 (Onboarding Detail) for cross-vertical lookup.');
    }

    // OMP rows
    var ompCompleted = ompRows.filter(function(r) { return r.status === 'COMPLETED'; });
    Logger.log('OMP total: ' + ompRows.length + '  COMPLETED: ' + ompCompleted.length);

    // Overlap by name
    var existingCount = ompCompleted.filter(function(r) {
      return r.name && existingNames[r.name.trim().toLowerCase()];
    }).length;
    Logger.log('OMP COMPLETED matched as Existing (by name): ' + existingCount);

    // Search for the specific vendor the user mentioned
    var testName = 'shriman narayan fiber udhog private limited';
    var testRows = rows.filter(function(r) { return r.name && r.name.trim().toLowerCase() === testName; });
    Logger.log('Rows for "SHRIMAN NARAYAN FIBER UDHOG PRIVATE LIMITED": ' + testRows.length);
    testRows.forEach(function(r) {
      Logger.log('  vertical=' + r.vertical + ' bizVertical="' + r.bizVertical + '" status=' + r.status + ' id=' + r.id);
    });
  });
  Logger.log('\n=== Done ===');
}

// Per-vertical KPIs + onboarded-by-category breakdown + detail rows.
function vStats(data) {
  var now = new Date(), weekAgo = new Date(now.getTime() - 7 * 86400000);
  var nowMs = now.getTime();
  var AGE_WARN_MS = 14 * 86400000, AGE_DUE_MS = 30 * 86400000;

  // Single pass: all status / GST / new-vs-old / material / TAT / GMV / aging counts.
  var total = data.length;
  var completed = 0, draft = 0, inReview = 0, rejected = 0;
  var withGST = 0, newVendors = 0, oldVendors = 0, completedThisWeek = 0;
  var hasTxnData = false, transactedCount = 0, txnCountSum = 0;
  var hasTxnValData = false, txnValSum = 0;
  var agingCount = 0, overdueCount = 0;
  var tats = [];
  var plasticTotal = 0, metalTotal = 0;
  var plasticOnboarded = 0, metalOnboarded = 0;
  var plasticTransacted = 0, metalTransacted = 0;
  var plasticNew = 0, metalNew = 0, plasticOld = 0, metalOld = 0;

  data.forEach(function(r) {
    var isDone = r.status === 'COMPLETED';
    var isOld  = !!r.isOldVendor, isNew = !r.isOldVendor;
    var cat    = String(r.category || '').trim().toLowerCase();
    var isPl   = cat === 'plastic', isMt = cat === 'metal';

    if (isDone)                    completed++;
    else if (r.status === 'DRAFT')      draft++;
    else if (r.status === 'IN_REVIEW')  inReview++;
    else if (r.status === 'REJECTED')   rejected++;

    if (isDone && r.hasGST)                                       withGST++;
    if (isDone && isNew)                                           newVendors++;
    if (isDone && isOld)                                           oldVendors++;
    if (isDone && r.onboardedDate && r.onboardedDate >= weekAgo)  completedThisWeek++;

    // onbTAT is the 5292-derived In-Review→now/completion TAT (see getTATLookup_),
    // so this average spans every record that has entered review — completed and
    // still-in-flight alike. TAT_MAX_DAYS clamps re-import/outlier noise.
    if (r.onbTAT !== null && r.onbTAT >= 0 && r.onbTAT <= TAT_MAX_DAYS) tats.push(r.onbTAT);

    if (r.hasTxn)        hasTxnData = true;
    if (r.hasTransacted) { transactedCount++; txnCountSum += (r.txnCount || 1); }
    if (r.txnValue !== null && r.txnValue !== undefined && r.txnValue > 0) {
      hasTxnValData = true; txnValSum += r.txnValue;
    }

    if (!isDone && r.status !== 'REJECTED' && r.createdDate) {
      var elapsed = nowMs - r.createdDate.getTime();
      if (elapsed > AGE_DUE_MS)       { overdueCount++; agingCount++; }
      else if (elapsed > AGE_WARN_MS) { agingCount++; }
    }

    if (isPl) {
      plasticTotal++;
      if (isDone)          plasticOnboarded++;
      if (r.hasTransacted) plasticTransacted++;
      if (isDone && isNew) plasticNew++;
      if (isDone && isOld) plasticOld++;
    }
    if (isMt) {
      metalTotal++;
      if (isDone)          metalOnboarded++;
      if (r.hasTransacted) metalTransacted++;
      if (isDone && isNew) metalNew++;
      if (isDone && isOld) metalOld++;
    }
  });

  var transacted    = hasTxnData    ? transactedCount : null;
  var totalTxnValue = hasTxnValData ? txnValSum       : null;

  var catMap = {};
  data.forEach(function(r) {
    var c = r.category || 'Others';
    if (!catMap[c]) catMap[c] = {
      category: c, total: 0, onboarded: 0,
      draft: 0, inReview: 0, rejected: 0,
      transacted: 0, hasTxn: false,
      vendorTypes: {},
    };
    var cs = catMap[c];
    cs.total++;
    if (r.status === 'COMPLETED')  cs.onboarded++;
    if (r.status === 'DRAFT')      cs.draft++;
    if (r.status === 'IN_REVIEW')  cs.inReview++;
    if (r.status === 'REJECTED')   cs.rejected++;
    if (r.hasTxn)                  cs.hasTxn = true;
    if (r.hasTransacted)           cs.transacted++;
    var vt = (r.vendorType || '').trim() || 'Unknown';
    cs.vendorTypes[vt] = (cs.vendorTypes[vt] || 0) + 1;
  });
  var categories = Object.keys(catMap).map(function(k) {
    var cat = catMap[k];
    var vtArr = Object.keys(cat.vendorTypes)
      .map(function(vt) { return { type: vt, count: cat.vendorTypes[vt] }; })
      .filter(function(x) { return x.type !== 'Unknown' && x.type; })
      .sort(function(a, b) { return b.count - a.count; });
    if (cat.vendorTypes['Unknown']) vtArr.push({ type: 'Unknown', count: cat.vendorTypes['Unknown'] });
    return {
      category: cat.category, total: cat.total, onboarded: cat.onboarded,
      draft: cat.draft, inReview: cat.inReview, rejected: cat.rejected,
      transacted: cat.transacted, hasTxn: cat.hasTxn,
      vendorTypes: vtArr,
    };
  }).sort(function(a, b) { return b.onboarded - a.onboarded || b.total - a.total; });

  // Per-financial-year breakdown (India FY: Apr 1 – Mar 31), newest FY first.
  var fyMap = {};
  data.forEach(function(r) {
    if (!r.createdDate) return;
    var y = r.createdDate.getFullYear(), m = r.createdDate.getMonth();
    var fyStart = m >= 3 ? y : y - 1;
    var fyKey   = 'FY ' + String(fyStart).slice(2) + '-' + String(fyStart + 1).slice(2);
    if (!fyMap[fyKey]) fyMap[fyKey] = { fyKey: fyKey, fyStart: fyStart, rows: [] };
    fyMap[fyKey].rows.push(r);
  });
  var fyBreakdown = Object.keys(fyMap).map(function(fk) {
    var d = fyMap[fk].rows;
    var fyTotal     = d.length;
    var fyCompleted = countFn(d, function(r) { return r.status === 'COMPLETED'; });
    var fyHasTxn    = d.some(function(r) { return r.hasTxn; });
    var fyTransacted = fyHasTxn ? countFn(d, function(r) { return r.hasTransacted; }) : null;
    var fyTxnSum = 0, fyHasTxnVal = false;
    d.forEach(function(r) {
      if (r.txnValue !== null && r.txnValue !== undefined && r.txnValue > 0) { fyHasTxnVal = true; fyTxnSum += r.txnValue; }
    });
    var fyWithGST = d.some(function(r) { return r.gstin; })
      ? countFn(d, function(r) { return r.status === 'COMPLETED' && r.hasGST; }) : null;
    var fyTats = d.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0 && r.onbTAT <= TAT_MAX_DAYS; })
                  .map(function(r) { return r.onbTAT; });
    return {
      fy:            fk,
      fyStart:       fyMap[fk].fyStart,
      total:         fyTotal,
      onboarded:     fyCompleted,
      completionPct: pct(fyCompleted, fyTotal),
      transacted:    fyTransacted,
      pctTransacted: fyTransacted === null ? null : pct(fyTransacted, fyCompleted),
      totalTxnValue: fyHasTxnVal ? fyTxnSum : null,
      withGST:       fyWithGST,
      avgTAT:        fyTats.length ? Math.round(avg(fyTats)) : null,
    };
  }).filter(function(f) { return f.fyStart >= 2019; })
    .sort(function(a, b) { return b.fyStart - a.fyStart; });

  return {
    total: total,
    completed: completed,
    draft: draft,
    inReview: inReview,
    rejected: rejected,
    withGST: withGST,
    completionPct: pct(completed, total),
    completedThisWeek: completedThisWeek,
    avgTAT: tats.length ? Math.round(avg(tats)) : null,
    transacted: transacted,
    pctTransacted: transacted === null ? null : pct(transacted, completed),
    totalTxnCount: hasTxnData ? txnCountSum : null,
    totalTxnValue: totalTxnValue,
    aging:       agingCount   > 0 ? agingCount   : null,
    overdue:     overdueCount > 0 ? overdueCount : null,
    newVendors: newVendors,
    oldVendors: oldVendors,
    plasticTotal: plasticTotal,       metalTotal: metalTotal,
    plasticOnboarded: plasticOnboarded, metalOnboarded: metalOnboarded,
    plasticTransacted: plasticTransacted, metalTransacted: metalTransacted,
    plasticNew: plasticNew,           metalNew: metalNew,
    plasticOld: plasticOld,           metalOld: metalOld,
    fyBreakdown: fyBreakdown,
    categories:  categories,
    rowsTotal: total,
  };
}

function vertRow(r) {
  return {
    id: r.id, name: r.name, category: r.category, vendorType: r.vendorType,
    status: r.status, gstin: r.gstin, hasGST: r.hasGST, state: r.state,
    createdDate: fmtDate(r.createdDate), onbDate: fmtDate(r.onboardedDate),
    tat: (r.onbTAT === null ? '—' : r.onbTAT),
    hasTxn:       r.hasTxn,
    hasTransacted: r.hasTransacted,
    txnCount:     r.txnCount || 1,
    txnValue:     r.txnValue !== null && r.txnValue !== undefined ? r.txnValue : null,
    txnDate:      fmtDate(r.txnDate),
    isOldVendor:  !!r.isOldVendor,
  };
}

// ─────────────────────────────────────────────────────────────
// DATE / PERIOD
// ─────────────────────────────────────────────────────────────
function applyDateFilter(r, f) {
  if (!f || !f.period || f.period === 'All') return true;
  var d = r.createdDate;
  if (!d) return false;
  var now = new Date(), ps = null, pe = null;
  if (f.period === 'Today') {
    ps = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    pe = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (f.period === 'ThisWeek') {
    // T-7: rolling last 7 days
    ps = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7);
    pe = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (f.period === 'ThisMonth') {
    // T-30: rolling last 30 days
    ps = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 30);
    pe = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (f.period === 'MTD') {
    // Month to date: 1st of current calendar month → today
    ps = new Date(now.getFullYear(), now.getMonth(), 1);
    pe = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (f.period === 'YTD') {
    // Financial year to date: April 1 of current Indian FY → today
    var fyS = (now.getMonth() >= 3) ? now.getFullYear() : now.getFullYear() - 1;
    ps = new Date(fyS, 3, 1);
    pe = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  } else if (f.period === 'Custom' && f.startDate && f.endDate) {
    ps = parseYMD(f.startDate, false);
    pe = parseYMD(f.endDate, true);
  } else if (String(f.period).indexOf('FY') === 0) {
    var m = String(f.period).match(/FY(\d{2})-(\d{2})/);
    if (m) { var y = 2000 + parseInt(m[1], 10); ps = new Date(y, 3, 1); pe = new Date(y + 1, 2, 31, 23, 59, 59, 999); }
  }
  if (ps && d < ps) return false;
  if (pe && d > pe) return false;
  return true;
}
function parseYMD(s, endOfDay) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s || '').trim());
  if (!m) return null;
  return endOfDay ? new Date(+m[1], +m[2] - 1, +m[3], 23, 59, 59, 999) : new Date(+m[1], +m[2] - 1, +m[3]);
}
function fyStartYear(d) { return d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1; }
function buildFiscalYears(all) {
  var nowFY = fyStartYear(new Date()), minFY = nowFY;
  all.forEach(function(r) {
    var d = r.createdDate || r.onboardedDate;
    if (d) { var y = fyStartYear(d); if (y >= 2019 && y < minFY) minFY = y; }
  });
  var list = [];
  for (var y = nowFY; y >= minFY; y--) list.push('FY' + String(y).slice(2) + '-' + String(y + 1).slice(2));
  return list;
}

// ─────────────────────────────────────────────────────────────
// FIELD PARSERS / NORMALIZERS
// ─────────────────────────────────────────────────────────────
function buildIndex(headers) { var idx = {}; headers.forEach(function(h, i) { idx[h] = i; }); return idx; }
function gv(row, idx, key) { return (idx[key] !== undefined) ? row[idx[key]] : ''; }
// Returns the value from the first column key that actually exists in idx.
// Unlike || chaining, this preserves 0 and other falsy-but-valid cell values.
function firstDef_(row, idx, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (idx[keys[i]] !== undefined) return row[idx[keys[i]]];
  }
  return '';
}

function parseDate(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') { if (val < 1) return null; var dn = new Date((val - 25569) * 86400000); return isNaN(dn.getTime()) ? null : dn; }
  var s = String(val).trim();
  if (!s || s === '—' || s === '-' || s === 'N/A') return null;
  var noTime = s.replace(/,?\s+\d{1,2}:\d{2}(:\d{2})?(\s*(am|pm))?/i, '').trim();
  var d1 = new Date(noTime);
  if (!isNaN(d1.getTime())) return d1;
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}
function dateDiffDays(d1, d2) {
  if (!d1 || !d2) return null;
  var a = Date.UTC(d1.getFullYear(), d1.getMonth(), d1.getDate());
  var b = Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate());
  return Math.round((b - a) / 86400000);
}
function isValidGSTIN(g) { return /^[0-9A-Z]{15}$/.test(String(g || '').trim().toUpperCase()); }

function normStatus(v) {
  var s = String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  if (!s) return 'UNKNOWN';
  var map = {
    COMPLETED: 'COMPLETED', COMPLETE: 'COMPLETED', ONBOARDED: 'COMPLETED', ACTIVE: 'COMPLETED',
    DRAFT: 'DRAFT', PENDING: 'DRAFT', NEW: 'DRAFT', INITIATED: 'DRAFT', INCOMPLETE: 'DRAFT',
    IN_REVIEW: 'IN_REVIEW', INREVIEW: 'IN_REVIEW', REVIEW: 'IN_REVIEW', UNDER_REVIEW: 'IN_REVIEW', IN_PROGRESS: 'IN_REVIEW', SUBMITTED: 'IN_REVIEW',
    REJECTED: 'REJECTED', REJECT: 'REJECTED', DECLINED: 'REJECTED',
  };
  return map[s] || s;
}
function normCategory(v) { var s = String(v || '').trim(); return s || 'Others'; }

function count(arr, field, val) { return arr.filter(function(r) { return r[field] === val; }).length; }
function countFn(arr, fn) { return arr.filter(fn).length; }
function avg(arr) { return arr.length ? arr.reduce(function(a, b) { return a + b; }, 0) / arr.length : 0; }
function pct(n, d) { return d ? Math.round((n / d) * 100) : 0; }
function fmtDate(d) {
  if (!d) return '—';
  var dd = d.getDate(), mm = d.getMonth() + 1, yy = d.getFullYear();
  return (dd < 10 ? '0' : '') + dd + '/' + (mm < 10 ? '0' : '') + mm + '/' + yy;
}
function slug(s) { return String(s).replace(/[^A-Za-z0-9]+/g, '') || 'V'; }
function codeFor(name) {
  var s = String(name || '').trim();
  var words = s.split(/[\s/&-]+/).filter(Boolean);
  if (words.length > 1) return words.map(function(w) { return w.charAt(0); }).join('').slice(0, 4).toUpperCase();
  return s.slice(0, 3).toUpperCase();
}

// ═══════════════════════════════════════════════════════════════
// GST PAYABLES — Vendor Payables Register
// ───────────────────────────────────────────────────────────────
// A second, independent dataset surfaced through its own dashboard tab.
// The first tab ("Vendor Payables") of the GST Payables spreadsheet is
// read server-side by ID (the deploying account's access — the sheet can
// stay private) and aggregated into a payables control view: outstanding
// KPIs, aging-by-financial-year, GST-registration mix, vertical exposure,
// compliance rates, and the full vendor register.
//
// Column names are resolved through a variant table (gpField_) so the
// dashboard keeps working even if headers differ slightly from what the
// sheet ships today. Run diagnoseGSTPayables() in the Apps Script editor
// to print the sheet's actual headers and confirm the mapping.
// ═══════════════════════════════════════════════════════════════
var GST_PAYABLES = {
  SHEET_ID: '1rhr8Omd99OrN3c1h1uu6VvTa-aolAmRY58p19NDlz54',
  CACHE_TTL: 300,   // 5-minute cache, same cadence as the onboarding views
};

// Candidate header names (normalized: lowercased, spaces→_, symbols stripped)
// for each logical field. First match in the sheet wins.
var GST_FIELDS = {
  name:      ['vendor', 'vendor_name', 'party', 'party_name', 'supplier', 'supplier_name', 'name', 'legal_name', 'trade_name', 'company', 'company_name'],
  gstNo:     ['gst_no', 'gstin', 'gst_number', 'gstno', 'gst', 'gstin_no', 'gst_in'],
  gstStatus: ['gst_status', 'gstin_status', 'registration_status', 'gst_registration_status', 'status', 'gst_state_status'],
  vertical:  ['vertical', 'business_vertical', 'category', 'business_category', 'bu', 'business_unit', 'segment'],
  msme:      ['msme', 'msme_status', 'msme_category', 'msme_type', 'msme_classification', 'msme_registration'],
  total:     ['total_balance', 'total_outstanding', 'outstanding', 'total_payable', 'closing_balance', 'net_balance', 'total_due', 'grand_total', 'payable', 'balance', 'total'],
  material:  ['material_balance', 'material', 'goods_balance', 'goods_component', 'material_outstanding', 'material_value', 'material_amount', 'basic_balance', 'basic_amount', 'taxable_value'],
  gst:       ['gst_balance', 'gst_component', 'tax_balance', 'tax_component', 'gst_amount', 'gst_value', 'gst_outstanding', 'gst_payable', 'tax_amount', 'tax_value'],
  fy:        ['financial_year', 'fy', 'year', 'invoice_fy', 'ageing_year', 'aging_year', 'fin_year', 'fy_year', 'financial_yr', 'invoice_year'],
  sda:       ['sda_signed', 'sda', 'sda_status'],
  ledger:    ['ledger_reconciled', 'ledger_reconciliation', 'ledger', 'reconciled', 'reconciliation', 'ledger_status'],
  osv:       ['osv_positive', 'osv', 'osv_status'],
  itcClosed: ['itc_closed', 'itc_status', 'itc', 'itc_closure', 'itc_reversal'],
  itcCheck:  ['itc_check', 'itc_eligible', 'itc_eligibility', 'itc_available', 'itc_flag'],
  state:     ['state', 'gst_state', 'location', 'region'],
};

// First header name from `variants` that exists in the sheet, or null.
function gpField_(idx, variants) {
  for (var i = 0; i < variants.length; i++) {
    if (idx[variants[i]] !== undefined) return variants[i];
  }
  return null;
}
function gpVal_(row, idx, variants) {
  var key = gpField_(idx, variants);
  return key ? row[idx[key]] : '';
}

// Parse a money/number cell: strips ₹, commas, spaces, any non-numeric noise.
function gpNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// Interpret a compliance cell as a yes/no flag.
function gpBool_(v) {
  var s = String(v || '').trim().toLowerCase();
  if (!s) return false;
  return ({ 'yes': 1, 'y': 1, 'true': 1, '1': 1, 'done': 1, 'closed': 1, 'signed': 1,
            'positive': 1, 'reconciled': 1, 'active': 1, 'complete': 1, 'completed': 1,
            'ok': 1, 'cleared': 1, 'yes ': 1, '✓': 1, 'true ': 1 })[s] === 1;
}

// Resolve the target tab: prefer the one literally named "Vendor Payables"
// (case/spacing-insensitive), else fall back to the first tab. Guards against
// a reordered workbook or a hidden helper sheet sitting at index 0.
function gpOpenSheet_() {
  var ss = SpreadsheetApp.openById(GST_PAYABLES.SHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var nm = String(sheets[i].getName() || '').trim().toLowerCase().replace(/\s+/g, ' ');
    if (nm === 'vendor payables' || nm === 'vendor payable' || nm === 'vendorpayables') return sheets[i];
  }
  return sheets[0];
}

// GST-specific sheet reader that auto-detects the header row.
// Strategy: skip rows where every non-empty cell is a number (title/total rows);
// use the first row that has at least one text cell and more text cells than
// numeric cells. This reliably handles sheets with a report-title row above
// the actual column header row (headers in row 2, data rows 3+).
function gpReadSheet_(sheet) {
  var vals = sheet.getDataRange().getValues();
  if (!vals.length) return { headers: [], rawHeaders: [], rows: [], headerRowIdx: 0 };

  var headerRowIdx = 0;
  var limit = Math.min(15, vals.length);
  for (var i = 0; i < limit; i++) {
    var textCnt = 0, numCnt = 0;
    vals[i].forEach(function(cell) {
      if (cell === '' || cell === null || cell === undefined) return;
      var s = String(cell).trim();
      if (!s) return;
      // Treat Dates and pure numbers as "numeric"
      if (cell instanceof Date || typeof cell === 'number' ||
          (s !== '' && !isNaN(parseFloat(s)) && isFinite(s.replace(/,/g, '')))) {
        numCnt++;
      } else {
        textCnt++;
      }
    });
    // First row with at least one text cell and more text than numbers → header row
    if (textCnt > 0 && textCnt >= numCnt) {
      headerRowIdx = i;
      break;
    }
  }

  var rawCells = vals[headerRowIdx];
  var rawHeaders = rawCells.map(function(h) { return String(h || '').trim(); });
  var headers = rawHeaders.map(function(h) {
    return h.toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  var rows = vals.slice(headerRowIdx + 1).filter(function(row) {
    return row.some(function(c) { return c !== '' && c !== null && c !== undefined; });
  });
  return { headers: headers, rawHeaders: rawHeaders, rows: rows, headerRowIdx: headerRowIdx };
}

// Read + normalize the vendor rows from the GST Payables sheet.
// Returns the vendors plus diagnostics (tab name, raw row count, headers,
// resolved field mapping) so the UI can distinguish a truly empty sheet from
// a header-mapping miss, and surface the real column names for verification.
function gpReadVendors_() {
  var sheet = gpOpenSheet_();
  var raw = gpReadSheet_(sheet);
  var idx = buildIndex(raw.headers);
  var mapping = {};
  Object.keys(GST_FIELDS).forEach(function(k) { mapping[k] = gpField_(idx, GST_FIELDS[k]) || null; });

  // Match only the exact financial-year columns present in the sheet.
  // Normalised headers strip the dash (2021-22 → 202122) so we match rawHeaders.
  var YEAR_HEADERS = {
    '2021-22': 'FY 21-22', '2022-23': 'FY 22-23', '2023-24': 'FY 23-24',
    '2024-25': 'FY 24-25', '2025-26': 'FY 25-26', '2026-27': 'FY 26-27',
  };
  var yearCols = [];
  raw.rawHeaders.forEach(function(rh, ci) {
    var key = String(rh || '').trim();
    if (YEAR_HEADERS[key]) yearCols.push({ colIdx: ci, fy: YEAR_HEADERS[key] });
  });

  var vendors = raw.rows.map(function(row) {
    var total = gpNum_(gpVal_(row, idx, GST_FIELDS.total));
    var hasMat = gpField_(idx, GST_FIELDS.material) !== null;
    var hasGst = gpField_(idx, GST_FIELDS.gst) !== null;
    var mat = hasMat ? gpNum_(gpVal_(row, idx, GST_FIELDS.material)) : null;
    var gst = hasGst ? gpNum_(gpVal_(row, idx, GST_FIELDS.gst)) : null;
    if (mat === null && gst !== null) mat = total - gst;
    if (gst === null && mat !== null) gst = total - mat;
    if (!total && (mat || gst)) total = (mat || 0) + (gst || 0);

    // Per-year outstanding from the year columns. 0 values are omitted.
    var yearAmounts = {};
    yearCols.forEach(function(yc) {
      var v = gpNum_(row[yc.colIdx]);
      if (v) yearAmounts[yc.fy] = v;
    });

    return {
      name:        String(gpVal_(row, idx, GST_FIELDS.name) || '').trim(),
      gstNo:       String(gpVal_(row, idx, GST_FIELDS.gstNo) || '').trim(),
      gstStatus:   String(gpVal_(row, idx, GST_FIELDS.gstStatus) || '').trim() || 'Unknown',
      vertical:    String(gpVal_(row, idx, GST_FIELDS.vertical) || '').trim() || 'Others',
      msme:        String(gpVal_(row, idx, GST_FIELDS.msme) || '').trim() || 'Not Registered',
      state:       String(gpVal_(row, idx, GST_FIELDS.state) || '').trim(),
      total:       total,
      material:    mat || 0,
      gst:         gst || 0,
      yearAmounts: yearAmounts,
      sda:         gpBool_(gpVal_(row, idx, GST_FIELDS.sda)),
      ledger:      gpBool_(gpVal_(row, idx, GST_FIELDS.ledger)),
      osv:         gpBool_(gpVal_(row, idx, GST_FIELDS.osv)),
      itcClosed:   gpBool_(gpVal_(row, idx, GST_FIELDS.itcClosed)),
      itcCheck:    String(gpVal_(row, idx, GST_FIELDS.itcCheck) || '').trim(),
    };
  }).filter(function(r) { return r.name || r.gstNo || r.total > 0; });
  return { vendors: vendors, yearCols: yearCols, headers: raw.headers, rawHeaders: raw.rawHeaders, rawRowCount: raw.rows.length, headerRowIdx: raw.headerRowIdx, tab: sheet.getName(), mapping: mapping };
}

// Entry point — aggregated GST Payables dashboard for the requested FY filter.
function getGSTPayablesData(filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var year = f.year || 'All';

    var cacheKey = 'gstpay_v10_' + JSON.stringify([year]);
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    var read = gpReadVendors_();
    var all = read.vendors;

    // FY breakdown: sum each year column across all vendors (unfiltered).
    // Material/GST per year derived from each vendor's overall ratio.
    var fyBreakMap = {};
    all.forEach(function(r) {
      var matRatio = r.total > 0 ? r.material / r.total : 0;
      var gstRatio = r.total > 0 ? r.gst      / r.total : 0;
      Object.keys(r.yearAmounts || {}).forEach(function(fy) {
        var amt = r.yearAmounts[fy] || 0;
        if (!amt) return;
        if (!fyBreakMap[fy]) fyBreakMap[fy] = { total: 0, material: 0, gst: 0, vendors: 0 };
        fyBreakMap[fy].total    += amt;
        fyBreakMap[fy].material += amt * matRatio;
        fyBreakMap[fy].gst      += amt * gstRatio;
        fyBreakMap[fy].vendors  += 1;
      });
    });
    var fyList = Object.keys(fyBreakMap).sort();

    // Year filter: include vendors with a non-zero amount in the selected year,
    // and replace total/material/gst with that year's specific amounts.
    var rows;
    if (year === 'All') {
      rows = all;
    } else {
      rows = all.filter(function(r) { return (r.yearAmounts || {})[year] > 0; })
        .map(function(r) {
          var yearAmt  = r.yearAmounts[year] || 0;
          var matRatio = r.total > 0 ? r.material / r.total : 0;
          var gstRatio = r.total > 0 ? r.gst      / r.total : 0;
          var matAmt = Math.round(yearAmt * matRatio);
          return {
            name: r.name, gstNo: r.gstNo, gstStatus: r.gstStatus,
            vertical: r.vertical, msme: r.msme, state: r.state,
            total:    yearAmt,
            material: matAmt,
            gst:      yearAmt - matAmt,
            sda: r.sda, ledger: r.ledger, osv: r.osv,
            itcClosed: r.itcClosed, itcCheck: r.itcCheck,
          };
        });
    }

    var totalOut = 0, matSum = 0, gstSum = 0;
    var statusMap = {}, vertMap = {}, msmeMap = {};
    var comp = { sda: 0, ledger: 0, osv: 0, itcClosed: 0 };
    rows.forEach(function(r) {
      totalOut += r.total; matSum += r.material; gstSum += r.gst;
      statusMap[r.gstStatus] = (statusMap[r.gstStatus] || 0) + 1;
      vertMap[r.vertical]    = (vertMap[r.vertical] || 0) + r.total;
      msmeMap[r.msme]        = (msmeMap[r.msme] || 0) + 1;
      if (r.sda) comp.sda++;
      if (r.ledger) comp.ledger++;
      if (r.osv) comp.osv++;
      if (r.itcClosed) comp.itcClosed++;
    });
    var n = rows.length;

    function toList(map, key) {
      return Object.keys(map).map(function(k) { var o = { label: k }; o[key] = map[k]; return o; })
        .sort(function(a, b) { return b[key] - a[key]; });
    }

    var out = {
      success: true,
      lastUpdated: new Date().toISOString(),
      year: year,
      fyList: fyList,
      totalVendors: all.length,   // unfiltered — lets the UI tell "sheet empty" from "year filter empty"
      vendorCount: n,
      totalOutstanding: totalOut,
      materialBalance: matSum,
      gstBalance: gstSum,
      materialShare: totalOut ? Math.round(matSum / totalOut * 100) : 0,
      gstShare:      totalOut ? Math.round(gstSum / totalOut * 100) : 0,
      aging:         fyList.map(function(fy) { return { fy: fy, value: fyBreakMap[fy].total }; }),
      fyBreakdown:   fyList.map(function(fy) {
        var b = fyBreakMap[fy];
        var bTotal = Math.round(b.total);
        var bMat   = Math.round(b.material);
        return { fy: fy, total: bTotal, material: bMat, gst: bTotal - bMat, vendors: b.vendors };
      }),
      gstStatus:  toList(statusMap, 'count'),
      verticals:  toList(vertMap, 'value'),
      msme:       toList(msmeMap, 'count'),
      compliance: [
        { label: 'SDA Signed',        count: comp.sda,       pct: pct(comp.sda, n) },
        { label: 'Ledger Reconciled', count: comp.ledger,    pct: pct(comp.ledger, n) },
        { label: 'OSV Positive',      count: comp.osv,       pct: pct(comp.osv, n) },
        { label: 'ITC Closed',        count: comp.itcClosed, pct: pct(comp.itcClosed, n) },
      ],
      rows: rows.sort(function(a, b) { return b.total - a.total; }),
      // Diagnostics — lets the UI (and you) see the tab read, how many raw rows
      // it held, the actual column headers, and how each mapped. Column names
      // only; no row data leaks here.
      debug: { tab: read.tab, rawRows: read.rawRowCount, headers: read.headers, rawHeaders: read.rawHeaders, headerRowIdx: read.headerRowIdx, mapping: read.mapping, yearCols: (read.yearCols || []).map(function(c) { return c.fy; }) },
    };

    var s = JSON.stringify(out);
    try { cache.put(cacheKey, s, GST_PAYABLES.CACHE_TTL); } catch (e) { /* >100KB — recompute next time */ }
    return s;
  } catch (err) {
    return JSON.stringify({ success: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// Diagnostic — run from the Apps Script editor to print the sheet's real
// headers and a sample row, so the GST_FIELDS mapping can be verified.
function diagnoseGSTPayables() {
  try {
    var ss = SpreadsheetApp.openById(GST_PAYABLES.SHEET_ID);
    var sheet = gpOpenSheet_();
    var raw = gpReadSheet_(sheet);
    var idx = buildIndex(raw.headers);
    Logger.log('Spreadsheet: ' + ss.getName());
    Logger.log('Tab read:    ' + sheet.getName());
    Logger.log('All tabs:    ' + ss.getSheets().map(function(s) { return s.getName(); }).join(' | '));
    Logger.log('Data rows:   ' + raw.rows.length);
    Logger.log('Headers (' + raw.headers.length + '): ' + raw.headers.join(' | '));
    Logger.log('── Resolved field mapping ──');
    Object.keys(GST_FIELDS).forEach(function(k) {
      Logger.log('  ' + k + ' → ' + (gpField_(idx, GST_FIELDS[k]) || '‹unmapped›'));
    });
    if (raw.rows.length) Logger.log('Sample row: ' + JSON.stringify(raw.rows[0]));
  } catch (e) {
    Logger.log('❌ diagnoseGSTPayables failed: ' + e.message
      + '\n(If this is a permissions error, share the sheet with the deploying account.)');
  }
}


// ════════════════════════════════════════════════════════════════
// QUALITY DATA — Vendor Rating · OSV Status · Doc Completeness
// Reads from the three visible sheets created by Metabase.gs
// (syncRating / syncOSV / syncDocs).  Returns audience-split data
// { seller, buyer, combined } so each pipeline card shows its own
// quality metrics.  Falls back to combined if no split is possible.
// ════════════════════════════════════════════════════════════════
function getQualityData() {
  var CACHE_KEY = 'quality_data_v7';
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY);
  if (cached) return cached;

  // Build vendor-ID lookup sets from the hidden main sheets
  var sellerIds = buildQualityIdSet_('_mb_sellers', ['seller_id','id','vendor_id','vendorid']);
  var buyerIds  = buildQualityIdSet_('_mb_buyers',  ['buyer_id', 'id','vendor_id','vendorid']);
  var hasSplit  = Object.keys(sellerIds).length > 0 || Object.keys(buyerIds).length > 0;

  // OMP-onboarded map keyed by ID (for OSV matching)
  var ompMap   = buildOmpOnboardedMap_();
  var ompTotal = Object.keys(ompMap).length;

  // OMP-onboarded maps keyed by GSTIN (for Vendor Rating matching)
  var ompGstinSellers = buildOmpGstinMap_('_mb_sellers', 'seller');
  var ompGstinBuyers  = buildOmpGstinMap_('_mb_buyers',  'buyer');
  // Merged map — seller wins on GSTIN collision
  var ompGstinMap = {};
  Object.keys(ompGstinBuyers).forEach(function(g)  { ompGstinMap[g] = ompGstinBuyers[g]; });
  Object.keys(ompGstinSellers).forEach(function(g) { ompGstinMap[g] = ompGstinSellers[g]; });

  function mkR()  { return { ws: 0, rated: 0, total: 0, dist: [0,0,0,0,0] }; }
  function mkO()  { return { completed: 0, pending: 0, failed: 0, notInitiated: 0, total: 0 }; }
  function mkD()  { return { complete: 0, partial: 0, incomplete: 0, missing: 0, total: 0 }; }
  function mkAcc(){ return { r: mkR(), o: mkO(), d: mkD() }; }
  var acc = { seller: mkAcc(), buyer: mkAcc(), combined: mkAcc() };

  // Determine audience for a quality-sheet row
  function resolveAud_(row, headers, idCol, audCol) {
    if (audCol >= 0) {
      var av = String(row[audCol] || '').trim().toLowerCase();
      if (av.indexOf('sell') >= 0 || av === 's') return 'seller';
      if (av.indexOf('buy')  >= 0 || av === 'b') return 'buyer';
    }
    if (idCol >= 0 && hasSplit) {
      var vid = String(row[idCol] || '').trim();
      if (vid && sellerIds[vid]) return 'seller';
      if (vid && buyerIds[vid])  return 'buyer';
    }
    return 'combined';
  }
  function tgts_(aud) { return aud === 'combined' ? ['combined'] : [aud, 'combined']; }

  // ── Single pass: read "Vendor Rating" sheet for all three metrics
  // Col AE (index 30) = seller_rating (float, 1-5)
  // Col AF (index 31) = osv_consent   (CONSENT_ACCEPTED / CONSENT_PENDING / etc.)
  // Cols 32-42        = per-document 0/1 flags (11 documents)
  try {
    var vrd = qualityReadSheet_('Vendor Rating');
    if (vrd.rows.length) {
      var vrh = vrd.headers;

      var vidC  = qualityFindCol_(vrh, ['seller_id','id','vendor_id','vendorid']);
      var vauC  = qualityFindCol_(vrh, ['audience','type','aud','seller_buyer']);
      var nameC = qualityFindCol_(vrh, ['seller_name','name','vendor_name','business_name','company_name']);
      if (nameC < 0) nameC = 1;

      // GSTIN column for rating GSTIN-match lookup
      var vrGstC = qualityFindCol_(vrh, ['gstin','gst_number','gst_no','gstin_number','gst']);

      // Rating: seller_rating is col AE; positional fallback index 30
      var rC = qualityFindCol_(vrh, ['seller_rating','rating','vendor_rating','average_rating','avg_rating','score','overall_rating','stars']);
      if (rC < 0) rC = 30;

      // OSV: osv_consent is col AF; positional fallback index 31
      var oC = qualityFindCol_(vrh, ['osv_consent','osv_status','osv','site_visit_status','field_verification_status','verification_status']);
      if (oC < 0) oC = 31;

      // Doc columns: 11 individual document flags (present from col 32 onward)
      var DOC_NAMES = [
        'additional_documents','proof_of_premises','electricity_bill',
        'kyc_document','gst_portal_screenshot_bank_details','cancelled_cheque',
        'msme_certificate','aadhaar','owner_pan','entity_pan','gst_certificate'
      ];
      var DOC_LABELS = [
        'Additional Docs','Proof of Premises','Electricity Bill',
        'KYC Document','GST Screenshot','Cancelled Cheque',
        'MSME Certificate','Aadhaar','Owner PAN','Entity PAN','GST Certificate'
      ];
      var docIdx = DOC_NAMES.map(function(n) { return qualityFindCol_(vrh, [n]); });
      // Positional fallback: use indices 32-42 if none found by name
      if (docIdx.every(function(i) { return i < 0; })) {
        docIdx = [32,33,34,35,36,37,38,39,40,41,42];
      }

      var vendorRatings = [];
      var vendorOSV     = [];
      var vendorDocs    = [];

      vrd.rows.forEach(function(row) {
        var rowAud = resolveAud_(row, vrh, vidC, vauC);
        var ts = tgts_(rowAud);

        // — Rating (OMP-onboarded only, matched by GSTIN) —
        var vrGstin   = vrGstC >= 0 ? String(row[vrGstC] || '').trim() : '';
        var ompRInfo  = vrGstin ? ompGstinMap[vrGstin] : null;
        if (ompRInfo) {
          var rv   = parseFloat(row[rC]);
          var rAud = ompRInfo.aud;
          ['seller' === rAud ? 'seller' : 'buyer', 'combined'].forEach(function(t) {
            if (!isNaN(rv) && rv > 0 && rv <= 5) {
              acc[t].r.ws    += rv;
              acc[t].r.rated += 1;
              acc[t].r.dist[Math.min(4, Math.max(0, Math.round(rv) - 1))] += 1;
            }
          });
          if (!isNaN(rv) && rv > 0 && rv <= 5) {
            vendorRatings.push({
              id:               ompRInfo.id   || String(row[vidC] || '').trim().slice(0, 30),
              name:             ompRInfo.name || String(row[nameC] || '').trim().slice(0, 60),
              gstin:            vrGstin.slice(0, 20),
              rating:           Math.round(rv * 10) / 10,
              onboardingStatus: ompRInfo.onboardingStatus,
              category:         ompRInfo.category,
              aud:              rAud
            });
          }
        }

        // — OSV consent (OMP-onboarded sellers only; binary: verified vs not_initiated) —
        var vid_ = String(row[vidC] || '').trim();
        var ompInfo = ompMap[vid_];
        if (ompInfo) {
          var osRaw = String(row[oC] || '').trim();
          var osUp  = osRaw.toUpperCase().replace(/\s+/g, '_');
          var osvStatus = (osUp === 'CONSENT_ACCEPTED' || osUp === 'YES' || osUp === 'Y' || osUp === 'TRUE')
            ? 'verified' : 'not_initiated';
          if (osvStatus === 'verified') {
            acc['seller'].o.completed   += 1;
            acc['combined'].o.completed += 1;
          }
          vendorOSV.push({
            id:               vid_.slice(0, 30),
            name:             ompInfo.name || String(row[nameC] || '').trim().slice(0, 60),
            gstin:            ompInfo.gstin,
            consentRaw:       osRaw.slice(0, 30),
            status:           osvStatus,
            onboardingStatus: ompInfo.onboardingStatus,
            category:         ompInfo.category,
            aud:              'seller'
          });
        }

        // — Doc completeness: count how many of 11 docs are submitted (value > 0) —
        var submitted = 0, totalDocs = 0, missingDocNames = [];
        docIdx.forEach(function(didx, di) {
          if (didx < 0 || didx >= row.length) return;
          totalDocs++;
          var dv = parseInt(row[didx], 10);
          if (!isNaN(dv) && dv > 0) submitted++;
          else missingDocNames.push(DOC_LABELS[di]);
        });
        ts.forEach(function(t) {
          acc[t].d.total += 1;
          if      (totalDocs > 0 && submitted === totalDocs) acc[t].d.complete   += 1;
          else if (submitted > 0)                            acc[t].d.partial    += 1;
          else                                               acc[t].d.incomplete += 1;
        });
        vendorDocs.push({
          id:          String(row[vidC] || '').trim().slice(0, 30),
          name:        String(row[nameC] || '').trim().slice(0, 60),
          submitted:   submitted,
          total:       totalDocs,
          missingDocs: missingDocNames,
          aud:         rowAud
        });
      });
    }
  } catch (e) { Logger.log('getQualityData vendor-rating: ' + e.message); }

  // Rating denominator = OMP-onboarded count per audience (GSTIN-matched)
  acc['seller'].r.total   = Object.keys(ompGstinSellers).length;
  acc['buyer'].r.total    = Object.keys(ompGstinBuyers).length;
  acc['combined'].r.total = Object.keys(ompGstinSellers).length + Object.keys(ompGstinBuyers).length;

  // OSV denominator = OMP-onboarded count; not_initiated fills the gap
  ['seller', 'combined'].forEach(function(t) {
    acc[t].o.total        = ompTotal;
    acc[t].o.notInitiated = Math.max(0, ompTotal - acc[t].o.completed);
    acc[t].o.pending      = 0;
    acc[t].o.failed       = 0;
  });

  // ── Finalize accumulators → output shape ─────────────────────
  function fR(r) { return { avg: r.rated>0?Math.round(r.ws/r.rated*10)/10:null, total:r.total, rated:r.rated, dist:r.dist }; }
  function fO(o) { return { completed:o.completed, pending:o.pending, failed:o.failed, notInitiated:o.notInitiated, total:o.total }; }
  function fD(d) { return { complete:d.complete, partial:d.partial, incomplete:d.incomplete, missing:d.missing, total:d.total }; }
  function fA(a) { return { rating:fR(a.r), osv:fO(a.o), docs:fD(a.d) }; }

  var combined = fA(acc.combined);
  var seller   = fA(acc.seller);
  var buyer    = fA(acc.buyer);

  // If audience-specific buckets are empty, fall back to combined
  if (seller.rating.total === 0 && seller.osv.total === 0 && seller.docs.total === 0)
    seller = JSON.parse(JSON.stringify(combined));
  if (buyer.rating.total === 0 && buyer.osv.total === 0 && buyer.docs.total === 0)
    buyer  = JSON.parse(JSON.stringify(combined));

  var result = { success: true, lastUpdated: new Date().toISOString(), combined: combined, seller: seller, buyer: buyer, vendorRatings: vendorRatings || [], vendorOSV: vendorOSV || [], vendorDocs: vendorDocs || [] };
  var out = JSON.stringify(result);
  try { cache.put(CACHE_KEY, out, 300); } catch (e) {}
  return out;
}

// Builds a { id: true } lookup set from a Metabase-synced sheet.
function buildQualityIdSet_(sheetName, idCols) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var maxCol  = Math.min(sheet.getLastColumn(), 40);
  var vals    = sheet.getRange(1, 1, sheet.getLastRow(), maxCol).getValues();
  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  var idCol = -1;
  for (var i = 0; i < idCols.length; i++) {
    var idx = headers.indexOf(idCols[i]);
    if (idx >= 0) { idCol = idx; break; }
  }
  if (idCol < 0) return {};
  var set = {};
  vals.slice(1).forEach(function(row) { var id = String(row[idCol]||'').trim(); if (id) set[id] = true; });
  return set;
}

// Returns { id: { name, gstin, category, onboardingStatus } } for OMP-onboarded sellers.
function buildOmpOnboardedMap_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('_mb_sellers');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var maxCol  = Math.min(sheet.getLastColumn(), 60);
  var vals    = sheet.getRange(1, 1, sheet.getLastRow(), maxCol).getValues();
  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  function fc(cands) { return qualityFindCol_(headers, cands); }
  var idC  = fc(['seller_id','id','vendor_id','vendorid']);
  var nmC  = fc(['seller_name','name','vendor_name','business_name','company_name']);
  var gstC = fc(['gstin','gst_number','gst_no','gstin_number','gst']);
  var catC = fc(['business_category','category','vertical_category','cat']);
  var bvC  = fc(['business_vertical','vertical','biz_vertical']);
  var stC  = fc(['onboarding_status','status','onboard_status']);
  if (idC < 0) return {};
  var map = {};
  vals.slice(1).forEach(function(row) {
    var id = String(row[idC] || '').trim(); if (!id) return;
    var bv = bvC >= 0 ? String(row[bvC] || '').trim().toLowerCase() : '';
    var st = stC >= 0 ? String(row[stC] || '').trim().toUpperCase()  : '';
    if (bv.indexOf('open marketplace') < 0 && bv.indexOf('open_marketplace') < 0 && bv !== 'omp') return;
    if (st !== 'COMPLETED') return;
    map[id] = {
      name:             nmC  >= 0 ? String(row[nmC]  || '').trim().slice(0, 60) : '',
      gstin:            gstC >= 0 ? String(row[gstC] || '').trim().slice(0, 20) : '',
      category:         catC >= 0 ? String(row[catC] || '').trim().slice(0, 60) : '',
      onboardingStatus: st
    };
  });
  return map;
}

// Returns { gstin: { id, name, category, onboardingStatus, aud } } for OMP-onboarded records.
// sheetName = '_mb_sellers' or '_mb_buyers'; aud = 'seller' or 'buyer'.
function buildOmpGstinMap_(sheetName, aud) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var maxCol  = Math.min(sheet.getLastColumn(), 60);
  var vals    = sheet.getRange(1, 1, sheet.getLastRow(), maxCol).getValues();
  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  function fc(cands) { return qualityFindCol_(headers, cands); }
  var idC  = fc(['seller_id','buyer_id','id','vendor_id','vendorid']);
  var nmC  = fc(['seller_name','buyer_name','name','vendor_name','business_name','company_name']);
  var gstC = fc(['gstin','gst_number','gst_no','gstin_number','gst']);
  var catC = fc(['business_category','category','vertical_category','cat']);
  var bvC  = fc(['business_vertical','vertical','biz_vertical']);
  var stC  = fc(['onboarding_status','status','onboard_status']);
  if (gstC < 0) return {};
  var map = {};
  vals.slice(1).forEach(function(row) {
    var gstin = String(row[gstC] || '').trim(); if (!gstin) return;
    var bv = bvC >= 0 ? String(row[bvC] || '').trim().toLowerCase() : '';
    var st = stC >= 0 ? String(row[stC] || '').trim().toUpperCase()  : '';
    if (bv.indexOf('open marketplace') < 0 && bv.indexOf('open_marketplace') < 0 && bv !== 'omp') return;
    if (st !== 'COMPLETED') return;
    map[gstin] = {
      id:               idC  >= 0 ? String(row[idC]  || '').trim().slice(0, 30) : '',
      name:             nmC  >= 0 ? String(row[nmC]  || '').trim().slice(0, 60) : '',
      category:         catC >= 0 ? String(row[catC] || '').trim().slice(0, 60) : '',
      onboardingStatus: st,
      aud:              aud
    };
  });
  return map;
}

function qualityFindCol_(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function qualityReadSheet_(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };
  var vals    = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  var rows = vals.slice(1).filter(function(row) {
    return row.some(function(c) { return c !== '' && c !== null && c !== undefined; });
  });
  return { headers: headers, rows: rows };
}
