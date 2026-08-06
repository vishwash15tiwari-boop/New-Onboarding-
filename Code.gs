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
  // from this tab's Level 1 date; see getLevel1Lookup_(). Run debugTAT()
  // to inspect column detection and match rate against the live tab.
  TAT_DETAIL_SHEET: '_mb_detail',
};

// Candidate column names in the 5292 detail tab (headers are normalized to
// lower_snake_case by readSheetObj_). First populated match wins per record.
var TAT_COLS = {
  // TAT start = latest Level 1 date.
  level1:    ['latest_level_1','latest_level_1_date','level_1_date','level_1_at','level_1_on',
              'level_1','level1_date','level1','l1_date','l1','level_1_approval_date',
              'level_1_review_date','level_1_updated_date','level_1_completed_date'],
  // Fallback start when no Level 1 date exists.
  created:   ['onboarding_created_date','created_date','created_at','created_on',
              'onboarding_created_at','registration_date','signup_date'],
  // TAT end = Onboarded date.
  completed: ['onboarding_completed_date','completed_date','completion_date','onboarded_date',
              'completed_at','onboarded_at','onboarding_updated_date','completed_on'],
  status:    ['onboarding_status','status','onboarding_state','state'],
  id:        ['seller_id','buyer_id','vendor_id','onboarding_id','id','entity_id','user_id','account_id'],
  name:      ['seller_name','buyer_name','vendor_name','business_name','name','entity_name'],
};

// Mandatory-document catalogue — shared by the Vendor Rating parse and the
// Doc Completeness parse so both produce the same document keys/labels.
// The first 11 are the historically fixed columns (positions 32-42 in the
// Vendor Rating sheet). PWM Certificate (buyer-mandatory) is detected by name
// only — see DOC_ALIASES — so the positional fallback stays tied to the 11.
var DOC_FIXED_COUNT = 11;
var DOC_NAMES = [
  'additional_documents','proof_of_premises','electricity_bill',
  'kyc_document','gst_portal_screenshot_bank_details','cancelled_cheque',
  'msme_certificate','aadhaar','owner_pan','entity_pan','gst_certificate',
  'pwm_certificate'
];
var DOC_LABELS = [
  'Additional Docs','Proof of Premises','Electricity Bill',
  'KYC Document','GST Screenshot','Cancelled Cheque',
  'MSME Certificate','Aadhaar','Owner PAN','Entity PAN','GST Certificate',
  'PWM Certificate'
];
// Alternate column spellings for documents whose header isn't the canonical key.
var DOC_ALIASES = {
  gst_certificate: ['gst_certificate','gst_cert','gstin_certificate','gst_registration_certificate',
                    'gst_certificate_url','gst_certificate_link','gst_certificate_document','gstcertificate'],
  entity_pan:      ['entity_pan','entity_pan_card','entity_pancard','entitypan','company_pan',
                    'business_pan','firm_pan','entity_pan_url','entity_pan_document'],
  pwm_certificate: ['pwm_certificate','pwm','pwm_cert','pwm_certificate_document','pwm_certificate_url',
                    'pwm_certificate_link','pwmcertificate','plastic_waste_management',
                    'plastic_waste_management_certificate','pwm_document']
};
// Token sets for a last-resort fuzzy match of the buyer-mandatory documents
// against whatever the source sheet names them (all tokens must appear).
var DOC_TOKENS = {
  gst_certificate: ['gst', 'cert'],
  entity_pan:      ['entity', 'pan'],
  pwm_certificate: ['pwm']
};
// Resolve a document's column index in a header list, honouring aliases.
function docColIndex_(headers, key) {
  return qualityFindCol_(headers, DOC_ALIASES[key] || [key]);
}
// Alias match first; then a token-based fuzzy match (only for keys in DOC_TOKENS),
// so e.g. "gst_certificate_status" or "pwm_cert_url" is still found.
function docColFuzzy_(headers, key) {
  var i = docColIndex_(headers, key);
  if (i >= 0) return i;
  var toks = DOC_TOKENS[key];
  if (!toks) return -1;
  for (var h = 0; h < headers.length; h++) {
    var hh = headers[h];
    var ok = true;
    for (var t = 0; t < toks.length; t++) { if (hh.indexOf(toks[t]) === -1) { ok = false; break; } }
    if (ok) return h;
  }
  return -1;
}
// A document cell counts as "submitted" when it holds a positive flag, an
// affirmative word, or any content (URL / date / id) — but not a negative.
function isDocSubmitted_(v) {
  if (v === '' || v === null || v === undefined) return false;
  if (typeof v === 'number') return v > 0;
  var s = String(v).trim().toLowerCase();
  if (!s) return false;
  if (s === '0' || s === 'no' || s === 'n' || s === 'false' || s === 'na' || s === 'n/a'
      || s === '-' || s === '—' || s === 'not submitted' || s === 'not_submitted'
      || s === 'pending' || s === 'missing' || s === 'not uploaded' || s === 'not_uploaded') return false;
  return true;
}

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
    var cacheKey  = 'dash_v35_' + audience + '_' + periodKey;
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
      if (v.length <= 100000) batchCache['vrows_v16_' + audience + '_' + vc.key + '_' + periodKey] = v;
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
// Fast path: compose from individual dash_v35_ cache entries when both are warm
// (they are pre-warmed by syncAllOnboarding for every period). Only falls through
// to the slow double-read when both individual caches are cold.
function getCombinedDashboard(filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var periodKey = JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cacheKey  = 'dash_v35_cmb_' + periodKey;
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    // Try to compose from pre-warmed individual caches (zero extra reads).
    var sIndKey = 'dash_v35_seller_' + periodKey;
    var bIndKey = 'dash_v35_buyer_'  + periodKey;
    var sInd = cache.get(sIndKey);
    var bInd = cache.get(bIndKey);
    if (sInd && bInd) {
      try {
        var sParsed = JSON.parse(sInd);
        var bParsed = JSON.parse(bInd);
        var sHasData = sParsed && sParsed.success && sParsed.verticals &&
          Object.keys(sParsed.verticals).some(function(k) { return (sParsed.verticals[k].total || 0) > 0; });
        if (sParsed && sParsed.success && bParsed && bParsed.success && sHasData) {
          var composed = JSON.stringify({
            success:        true,
            lastUpdated:    sParsed.mbSyncedAt || new Date().toISOString(),
            dataSource:     sParsed.dataSource || 'sheets',
            verticalConfig: sParsed.verticalConfig,
            seller: { verticals: sParsed.verticals, fiscalYears: sParsed.fiscalYears },
            buyer:  { verticals: bParsed.verticals, fiscalYears: bParsed.fiscalYears },
          });
          try { cache.put(cacheKey, composed, CONFIG.CACHE_TTL); } catch (e) {}
          return composed;
        }
      } catch (e) { /* fall through to full read */ }
    }

    // Full read — used only when both individual caches are cold.
    var sRaw  = readData('seller');
    var bRaw  = readData('buyer');
    var sCfg  = AUDIENCE_CFG.seller;
    var bCfg  = AUDIENCE_CFG.buyer;
    var sRows = normalizeRows(sRaw, sCfg);
    var bRows = normalizeRows(bRaw, bCfg);
    var sDash = buildDashboard('seller', sCfg, sRows, f);
    var bDash = buildDashboard('buyer',  bCfg, bRows, f);

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

    var cacheKey = 'vrows_v16_' + audience + '_' + vertKey + '_'
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
      var cacheKey = 'vrows_v16_' + aud + '_OMP_' + periodKey;
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
// LEVEL 1 LOOKUP — the latest "Level 1" date per record, read from Metabase
// card 5292 ("Onboarding Detail"), synced into the _mb_detail tab every 5 min.
//
// TAT is computed in normalizeRows as Level 1 → Onboarded. The Onboarded (end)
// and Created (fallback start) dates come from the MAIN seller/buyer card — the
// same trusted dates the dashboard has always used — so pulling the Level 1
// date here only ever moves the start LATER, which can only REDUCE the TAT and
// never inflate it beyond the legacy Created→Onboarded span. When no Level 1
// date matches a record, TAT is exactly the legacy Created→Onboarded value.
// Rebuilt on every request so values refresh when the 5292 sync updates.
// ─────────────────────────────────────────────────────────────
var _level1LookupCache = null;   // per-execution cache; each request re-evaluates the module

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

function getLevel1Lookup_() {
  if (_level1LookupCache) return _level1LookupCache;
  var lk = { byId: {}, byName: {}, loaded: false };
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TAT_DETAIL_SHEET);
    if (sheet && sheet.getLastRow() > 1) {
      var d   = readSheetObj_(sheet);
      var idx = buildIndex(d.headers);
      var nameIds = {};   // name → {id, ambiguous} — a name shared by >1 id can't be matched safely
      function foldLatest_(map, key, dt) {
        if (!map[key] || dt > map[key]) map[key] = dt;   // keep the LATEST Level 1 date
      }
      d.rows.forEach(function(row) {
        var l1 = parseDate(firstVal_(row, idx, TAT_COLS.level1));
        if (!l1) return;
        var id   = String(firstVal_(row, idx, TAT_COLS.id) || '').replace(/,/g, '').trim();
        var name = String(firstVal_(row, idx, TAT_COLS.name) || '').trim().toLowerCase();
        if (id)   foldLatest_(lk.byId, id, l1);
        if (name) foldLatest_(lk.byName, name, l1);
        if (name && id) {
          var ni = nameIds[name];
          if (!ni) nameIds[name] = { id: id, ambiguous: false };
          else if (ni.id !== id) ni.ambiguous = true;
        }
      });
      // Drop names shared by multiple distinct ids — matching them would pull a
      // Level 1 date from an unrelated vendor.
      Object.keys(nameIds).forEach(function(nm) { if (nameIds[nm].ambiguous) delete lk.byName[nm]; });
      lk.loaded = Object.keys(lk.byId).length + Object.keys(lk.byName).length > 0;
    }
  } catch (e) { Logger.log('Level 1 lookup build failed: ' + e.message); }
  _level1LookupCache = lk;
  return lk;
}

// Latest Level 1 date for a record (Date), matched by id first then name, or null.
function lookupLevel1_(id, name) {
  var lk = getLevel1Lookup_();
  if (id)   { var k = String(id).replace(/,/g, '').trim(); if (lk.byId[k])   return lk.byId[k]; }
  if (name) { var n = String(name).trim().toLowerCase();   if (lk.byName[n]) return lk.byName[n]; }
  return null;
}

// ─────────────────────────────────────────────────────────────
// TAT DIAGNOSTIC — run manually from the Apps Script editor to verify the
// 5292 tab is wired correctly. Logs the actual headers, which candidate
// column matched each TAT field, coverage counts, sample TATs, and the
// match rate against the live seller/buyer records. If a field shows
// "(none detected)", add the real header to TAT_COLS.
// ─────────────────────────────────────────────────────────────
function debugTAT() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TAT_DETAIL_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('❌ "' + CONFIG.TAT_DETAIL_SHEET + '" tab missing or empty — run syncDetail() first.');
    return;
  }
  var d   = readSheetObj_(sheet);
  var idx = buildIndex(d.headers);
  Logger.log('════ TAT DEBUG (card 5292 → ' + CONFIG.TAT_DETAIL_SHEET + ') ════');
  Logger.log('Detail rows: ' + d.rows.length);
  Logger.log('Headers (' + d.headers.length + '): ' + d.headers.join(', '));

  function detected(cat) {
    var list = TAT_COLS[cat];
    for (var i = 0; i < list.length; i++) if (idx[list[i]] !== undefined) return list[i];
    return '(none detected)';
  }
  ['level1', 'created', 'completed', 'status', 'id', 'name'].forEach(function(c) {
    Logger.log('  ' + c + ' column → ' + detected(c));
  });

  var nL1 = 0, nCr = 0, nOn = 0, nDone = 0;
  d.rows.forEach(function(row) {
    if (parseDate(firstVal_(row, idx, TAT_COLS.level1)))    nL1++;
    if (parseDate(firstVal_(row, idx, TAT_COLS.created)))   nCr++;
    if (parseDate(firstVal_(row, idx, TAT_COLS.completed))) nOn++;
    if (normStatus(firstVal_(row, idx, TAT_COLS.status)) === 'COMPLETED') nDone++;
  });
  Logger.log('Coverage — Level1:' + nL1 + '  Created:' + nCr + '  Onboarded:' + nOn + '  Completed-status:' + nDone);

  _level1LookupCache = null;                 // force a fresh build
  var lk = getLevel1Lookup_();
  Logger.log('Level 1 lookup — byId:' + Object.keys(lk.byId).length + '  byName:' + Object.keys(lk.byName).length);

  // Compare the live Avg TAT (Level1→Onboarded) against the Created→Onboarded
  // baseline so the reduction from using Level 1 is visible, plus the Level 1
  // match rate. If avg is high, check which columns were detected above.
  ['seller', 'buyer'].forEach(function(aud) {
    try {
      var cfg  = AUDIENCE_CFG[aud];
      var rows = normalizeRows(readData(aud), cfg);
      var n = 0, sumTat = 0, l1matched = 0, sumBase = 0, nBase = 0;
      rows.forEach(function(r) {
        if (lookupLevel1_(r.id, r.name)) l1matched++;
        if (r.onbTAT !== null && r.onbTAT >= 0 && r.onbTAT <= TAT_MAX_DAYS) { sumTat += r.onbTAT; n++; }
        if (r.status === 'COMPLETED' && r.createdDate && r.onboardedDate) {
          var b = dateDiffDays(r.createdDate, r.onboardedDate);
          if (b !== null && b >= 0 && b <= TAT_MAX_DAYS) { sumBase += b; nBase++; }
        }
      });
      Logger.log(aud + ': ' + rows.length + ' records · Level 1 matched ' + l1matched
        + ' · Avg TAT (Level1→Onb) ' + (n ? Math.round(sumTat / n) + 'd over ' + n : '—')
        + ' · baseline (Created→Onb) ' + (nBase ? Math.round(sumBase / nBase) + 'd over ' + nBase : '—'));
    } catch (e) { Logger.log(aud + ' match check failed: ' + e.message); }
  });
  Logger.log('════ END ════');
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
    // Timestamp when the case entered the IN_REVIEW stage. Also drives review-age
    // on Kanban cards ("Xd waiting" counts from here, not from created date).
    var reviewDate = parseDate(
      gv(row, idx, 'in_review_date')      ||
      gv(row, idx, 'in_review_at')        ||
      gv(row, idx, 'submitted_date')      ||
      gv(row, idx, 'submitted_at')        ||
      gv(row, idx, 'review_date')         ||
      gv(row, idx, 'review_at')           ||
      gv(row, idx, 'review_started_date') ||
      gv(row, idx, 'review_started_at')   ||
      gv(row, idx, 'level_2_date')        ||
      gv(row, idx, 'level2_date')         ||
      gv(row, idx, 'l2_date')            || ''
    );

    // TAT = In Review → Onboarded, for completed records only. Measuring from the
    // moment the case entered review reflects the time verification actually took,
    // excluding however long the vendor sat in draft before submitting.
    // The start must fall inside [created, onboarded] to be trusted. When no review
    // date is present the start falls back to the Level 1 date (card 5292), then to
    // Created, so TAT stays populated for feeds that carry no review timestamp.
    var level1   = lookupLevel1_(recId, recName);
    var inWindow = function(d) { return d && created && onboarded && d >= created && d <= onboarded; };
    var tatStart = inWindow(reviewDate) ? reviewDate
                 : inWindow(level1)     ? level1
                 : created;
    var tat = (status === 'COMPLETED' && tatStart && onboarded) ? dateDiffDays(tatStart, onboarded) : null;
    if (tat !== null && tat < 0) tat = null;

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
    // Number of transactions for this vendor, taken from the feed's total_orders
    // column when present. This is authoritative: it is the real order count, unlike
    // the pre-dedup row tally below which only sees rows this query happened to emit.
    var ordersRaw = gv(row, idx, 'total_orders');
    if (ordersRaw === '' || ordersRaw === null || ordersRaw === undefined) {
      var _oAlias = ['total_order', 'orders_count', 'order_count', 'no_of_orders',
                     'num_orders', 'total_transactions', 'transaction_count', 'txn_count'];
      for (var _o = 0; _o < _oAlias.length; _o++) {
        if (idx[_oAlias[_o]] !== undefined) { ordersRaw = row[idx[_oAlias[_o]]]; break; }
      }
    }
    var _ordN = (ordersRaw !== '' && ordersRaw !== null && ordersRaw !== undefined)
      ? parseInt(String(ordersRaw).replace(/[,\s]/g, ''), 10)
      : NaN;
    var totalOrders = isNaN(_ordN) ? null : _ordN;

    // Transacted = explicit positive status OR a positive transaction value:
    // GMV only exists once a transaction has happened, so GMV > 0 is itself
    // proof of a transaction even when no status column is present.
    var txnStatusPositive = ({ 'TRANSACTED': 1, 'YES': 1, 'Y': 1, 'TRUE': 1, '1': 1, 'DONE': 1, 'ACTIVE': 1, 'TRANSACTED YES': 1 })[txn] === 1;
    var transacted = txnStatusPositive || (txnVal !== null && txnVal > 0);

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
      totalOrders:   totalOrders,
      onbTAT:        tat,
      reviewDate:    reviewDate,
    };
  }).filter(function(r) { return r.id || r.name; });
  // Count pre-dedup rows per (id+vertical) as per-vendor transaction count.
  // Metabase emits one row per transaction via joins; this count captures that.
  var txnCounts = {};
  // total_orders, when the feed carries it, wins over the row tally. Tracked as a max
  // per key because dedup keeps only one row and the column may be blank on some of them.
  var orderMax = {};
  normalized.forEach(function(r) {
    if (!r.id) return;
    var key = r.id + '\x00' + r.vertical;
    txnCounts[key] = (txnCounts[key] || 0) + 1;
    if (r.totalOrders != null) {
      orderMax[key] = (orderMax[key] == null) ? r.totalOrders : Math.max(orderMax[key], r.totalOrders);
    }
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
    // Prefer the feed's total_orders; fall back to the pre-dedup row tally.
    r.txnCount = (orderMax[key] != null) ? orderMax[key] : (txnCounts[key] || 1);
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
    verticals[vc.key] = vStats(byVert[vc.key] || [], vc.key);
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

// ─── DIAGNOSTIC — one-shot audit of EVERY column pulled from EVERY source ────
// Run from the Apps Script editor (Run ▸ auditMetabaseColumns) then View ▸ Logs.
// Read-only. For each onboarding card it logs the live source, row count, the
// full header list, which key columns are present vs MISSING, and the raw
// business_vertical / category / region distributions the dashboard depends on;
// then the header lists of the TAT / rating / docs sheets. Use it to confirm
// exactly what data is being pulled and that no expected column is absent.
function auditMetabaseColumns() {
  Logger.log('════════ METABASE / SHEET COLUMN AUDIT ════════');

  ['seller', 'buyer'].forEach(function(aud) {
    var cfg = AUDIENCE_CFG[aud];
    Logger.log('\n──────── ' + aud.toUpperCase()
      + '  (card ' + (CONFIG.MB_CARDS && CONFIG.MB_CARDS[aud] || '?')
      + ' / sheet ' + (CONFIG.MB_LOCAL_SHEETS && CONFIG.MB_LOCAL_SHEETS[aud] || '?') + ') ────────');
    var raw;
    try { raw = readData(aud); } catch (e) { Logger.log('  ✗ readData failed: ' + e.message); return; }
    Logger.log('  source: ' + raw.source + '   rows: ' + raw.rows.length);
    Logger.log('  headers (' + raw.headers.length + '): ' + raw.headers.join(', '));

    var must = [cfg.idCol, cfg.nameCol, cfg.vendorCol, cfg.onbCol,
      'business_vertical', 'business_category', 'onboarding_status',
      'gst_number', 'gstin', 'gstin_status', 'transaction_activation_status',
      'transaction_date', 'txn_date', 'state', 'onboarding_created_date', 'in_review_date'];
    var present = [], missing = [];
    must.forEach(function(c) { (raw.headers.indexOf(c) >= 0 ? present : missing).push(c); });
    Logger.log('  ✓ present: ' + present.join(', '));
    Logger.log('  ✗ MISSING: ' + (missing.length ? missing.join(', ') : '(none)'));

    var rows = normalizeRows(raw, cfg);
    var bv = {}, cat = {}, reg = {};
    rows.forEach(function(r) {
      var b = String(r.bizVertical || '(blank)').trim() || '(blank)';   bv[b]  = (bv[b] || 0) + 1;
      var c = String(r.category || '(blank)').trim() || '(blank)';      cat[c] = (cat[c] || 0) + 1;
      var g = stateToRegion_(r.state) || '(unmapped)';                  reg[g] = (reg[g] || 0) + 1;
    });
    Logger.log('  raw business_vertical → ' + JSON.stringify(bv));
    Logger.log('  category → ' + JSON.stringify(cat));
    Logger.log('  region (state→zone) → ' + JSON.stringify(reg));
  });

  [['_mb_detail', 'TAT / Onboarding Detail (card 5292)'],
   ['Vendor Rating', 'Vendor Ratings (card 5662)'],
   ['Doc Completeness', 'Doc Completeness (card 5674)']].forEach(function(pair) {
    Logger.log('\n──────── ' + pair[0] + '  [' + pair[1] + '] ────────');
    var sh = null;
    try { sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pair[0]); } catch (e) {}
    if (!sh || sh.getLastRow() < 1) { Logger.log('  (sheet not found or empty)'); return; }
    var d = readSheetObj_(sh);
    Logger.log('  rows: ' + d.rows.length + '   headers (' + d.headers.length + '): ' + d.headers.join(', '));
  });

  Logger.log('\n════════ END AUDIT ════════');
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
// vertKey is used only to gate Existing-vs-New: isOldVendor (set in
// buildDashboard) is computed for OMP rows exclusively, so newVendors/
// oldVendors are only meaningful there — every other vertical returns null
// rather than a fabricated "100% new" split (see buildDashboard's
// existingNames comment for why the match is OMP-scoped).
function vStats(data, vertKey) {
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

  data.forEach(function(r) {
    var isDone = r.status === 'COMPLETED';
    var isOld  = !!r.isOldVendor, isNew = !r.isOldVendor;

    if (isDone)                    completed++;
    else if (r.status === 'DRAFT')      draft++;
    else if (r.status === 'IN_REVIEW')  inReview++;
    else if (r.status === 'REJECTED')   rejected++;

    if (isDone && r.hasGST)                                       withGST++;
    if (isDone && isNew)                                           newVendors++;
    if (isDone && isOld)                                           oldVendors++;
    if (isDone && r.onboardedDate && r.onboardedDate >= weekAgo)  completedThisWeek++;

    // onbTAT is the Level1→Onboarded TAT (Level 1 from card 5292; Onboarded/Created
    // from the main card — see normalizeRows), defined only for completed records.
    // TAT_MAX_DAYS clamps outlier noise.
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
      txnValue: 0, hasTxnVal: false,
      vendorTypes: {},
      tats: [],
    };
    var cs = catMap[c];
    cs.total++;
    // Category TAT uses the same guard as the headline average so the numbers agree.
    if (r.onbTAT !== null && r.onbTAT >= 0 && r.onbTAT <= TAT_MAX_DAYS) cs.tats.push(r.onbTAT);
    if (r.status === 'COMPLETED')  cs.onboarded++;
    if (r.status === 'DRAFT')      cs.draft++;
    if (r.status === 'IN_REVIEW')  cs.inReview++;
    if (r.status === 'REJECTED')   cs.rejected++;
    if (r.hasTxn)                  cs.hasTxn = true;
    if (r.hasTransacted)           cs.transacted++;
    if (r.txnValue !== null && r.txnValue !== undefined && r.txnValue > 0) {
      cs.txnValue += r.txnValue; cs.hasTxnVal = true;
    }
    // Vendor Type breakdown counts ONBOARDED vendors only (the donut reflects
    // the vendor-type mix of onboarded Sellers/Buyers, not the full pipeline).
    if (r.status === 'COMPLETED') {
      var vt = (r.vendorType || '').trim() || 'Unknown';
      cs.vendorTypes[vt] = (cs.vendorTypes[vt] || 0) + 1;
    }
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
      txnValue: cat.hasTxnVal ? cat.txnValue : null,
      vendorTypes: vtArr,
      // Average In Review → Onboarded TAT for this category (null when no completed
      // record in the category carries a usable TAT).
      avgTAT:   cat.tats.length ? Math.round(avg(cat.tats)) : null,
      tatCount: cat.tats.length,
    };
  }).sort(function(a, b) { return b.onboarded - a.onboarded || b.total - a.total; });

  // Region-wise bifurcation (zonal grouping of the vendor's state). Rows with
  // no resolvable state are skipped so shares reflect located vendors only.
  var regionMap = {};
  data.forEach(function(r) {
    var reg = stateToRegion_(r.state);
    if (!reg) return;
    if (!regionMap[reg]) regionMap[reg] = {
      region: reg, total: 0, onboarded: 0, draft: 0, inReview: 0,
      txnValue: 0, hasTxnVal: false,
    };
    var rs = regionMap[reg];
    rs.total++;
    if (r.status === 'COMPLETED') rs.onboarded++;
    if (r.status === 'DRAFT')     rs.draft++;
    if (r.status === 'IN_REVIEW') rs.inReview++;
    if (r.txnValue !== null && r.txnValue !== undefined && r.txnValue > 0) {
      rs.txnValue += r.txnValue; rs.hasTxnVal = true;
    }
  });
  var regions = Object.keys(regionMap).map(function(k) {
    var g = regionMap[k];
    return {
      region: g.region, total: g.total, onboarded: g.onboarded,
      draft: g.draft, inReview: g.inReview,
      txnValue: g.hasTxnVal ? g.txnValue : null,
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
    // Existing-vs-New is only ever computed for OMP (see comment above) — every
    // other vertical gets null here, not a misleading 100%-new split.
    newVendors: vertKey === 'OMP' ? newVendors : null,
    oldVendors: vertKey === 'OMP' ? oldVendors : null,
    fyBreakdown: fyBreakdown,
    categories:  categories,
    regions:     regions,
    rowsTotal: total,
  };
}

function vertRow(r) {
  return {
    id: r.id, name: r.name, category: r.category, vendorType: r.vendorType,
    status: r.status, gstin: r.gstin, hasGST: r.hasGST, state: r.state,
    createdDate: fmtDate(r.createdDate), onbDate: fmtDate(r.onboardedDate),
    reviewDate:  fmtDate(r.reviewDate),
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
// India Standard Time = UTC+5:30, no DST. Period filtering compares IST CALENDAR
// DAYS (yyyymmdd integers) rather than raw timestamps, so "Today" always means
// today's actual date (e.g. 28.07.2026) regardless of the server/script timezone
// — the previous logic used the server clock and could resolve "Today" to the
// wrong calendar day (showing "No data").
var IST_OFFSET_MS = (5 * 60 + 30) * 60 * 1000;
// Calendar day (yyyymmdd) of a UTC-millisecond instant, read in UTC.
function _ymdIntUTC_(ms) {
  var t = new Date(ms);
  return t.getUTCFullYear() * 10000 + (t.getUTCMonth() + 1) * 100 + t.getUTCDate();
}
// A record date → its IST calendar day (shift the absolute instant into IST).
function istDayNum_(d) { return _ymdIntUTC_(d.getTime() + IST_OFFSET_MS); }

function applyDateFilter(r, f) {
  if (!f || !f.period || f.period === 'All') return true;
  // Filter by the date the record reached its CURRENT stage, so period views
  // reflect day-wise activity: an onboarded record is matched on its onboarded
  // date (a case onboarded Today appears even if it was created earlier); records
  // still in the pipeline are matched on their created date.
  var d = (r.status === 'COMPLETED' && r.onboardedDate) ? r.onboardedDate : r.createdDate;
  if (!d) return false;
  var recDay = istDayNum_(d);

  // "Now" in IST → calendar parts used to build the period's day-range.
  var n = new Date(Date.now() + IST_OFFSET_MS);
  var Y = n.getUTCFullYear(), M = n.getUTCMonth(), D = n.getUTCDate();
  var todayDay = Y * 10000 + (M + 1) * 100 + D;
  // Day-of parts already in IST → build boundary days without re-shifting.
  function bDay(y, m, d0) { return _ymdIntUTC_(Date.UTC(y, m, d0)); }

  var psDay = null, peDay = null;
  if (f.period === 'Today') {
    psDay = peDay = todayDay;
  } else if (f.period === 'ThisWeek') {           // rolling last 7 days
    psDay = bDay(Y, M, D - 6); peDay = todayDay;
  } else if (f.period === 'ThisMonth') {          // rolling last 30 days
    psDay = bDay(Y, M, D - 29); peDay = todayDay;
  } else if (f.period === 'MTD') {                // month to date
    psDay = bDay(Y, M, 1); peDay = todayDay;
  } else if (f.period === 'YTD') {                // Indian FY to date (Apr 1 →)
    var fyS = (M >= 3) ? Y : Y - 1;
    psDay = bDay(fyS, 3, 1); peDay = todayDay;
  } else if (f.period === 'Custom' && f.startDate && f.endDate) {
    psDay = _ymdIntStr_(f.startDate); peDay = _ymdIntStr_(f.endDate);
  } else if (String(f.period).indexOf('FY') === 0) {
    var m = String(f.period).match(/FY(\d{2})-(\d{2})/);
    if (m) { var y = 2000 + parseInt(m[1], 10); psDay = y * 10000 + 401; peDay = (y + 1) * 10000 + 331; }
  }
  if (psDay !== null && recDay < psDay) return false;
  if (peDay !== null && recDay > peDay) return false;
  return true;
}
// "yyyy-mm-dd" → yyyymmdd integer.
function _ymdIntStr_(s) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(s || '').trim());
  return m ? (+m[1]) * 10000 + (+m[2]) * 100 + (+m[3]) : null;
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

// ── State → Region (two-way North / South split) ────────────────────
// The Business Insights "Regional Split" card buckets every vendor into just
// North or South. South = the five southern states and their UTs; North =
// every other state/UT. Handles full names (case- & punctuation-insensitive),
// common aliases, and 2-digit GST state codes. Empty state → '' (skipped, never
// counted); any non-empty value that isn't recognised as South falls to North.
var SOUTH_STATES = (function() {
  var s = {};
  ['andhrapradesh','ap','karnataka','ka','kerala','kl','tamilnadu','tn','telangana','ts','tg',
   'puducherry','pondicherry','py','lakshadweep','ld',
   'andamannicobar','andamanandnicobar','andamanandnicobarislands','an']
    .forEach(function(n) { s[n] = true; });
  return s;
}());
// GST 2-digit codes for the southern states/UTs.
var SOUTH_GST_CODES = {
  '28': true, '29': true, '31': true, '32': true, '33': true, '34': true, '35': true, '36': true, '37': true,
};
function stateToRegion_(state) {
  var raw = String(state || '').trim();
  if (!raw) return '';
  // GST state code, e.g. "27" or "33-Tamil Nadu" → take the leading digits.
  var digits = raw.match(/^\s*(\d{1,2})\b/);
  if (digits) {
    var code = digits[1].length === 1 ? '0' + digits[1] : digits[1];
    return SOUTH_GST_CODES[code] ? 'South' : 'North';
  }
  // Normalise to a lookup key: lowercase, strip everything but a-z.
  var key = raw.toLowerCase().replace(/&/g, 'and').replace(/[^a-z]/g, '');
  return SOUTH_STATES[key] ? 'South' : 'North';
}

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
  var CACHE_KEY = 'quality_data_v9';   // v9: buyer docs sourced from _mb_buyers
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

      // Doc columns: individual document flags (the fixed 11 sit at cols 32-42;
      // PWM is detected by name via docColIndex_ aliases).
      var docIdx = DOC_NAMES.map(function(n) { return docColIndex_(vrh, n); });
      // Positional fallback for the fixed 11 only when none were found by name;
      // never overwrites name-detected extras like PWM.
      if (docIdx.slice(0, DOC_FIXED_COUNT).every(function(i) { return i < 0; })) {
        for (var _pf = 0; _pf < DOC_FIXED_COUNT; _pf++) docIdx[_pf] = 32 + _pf;
      }
      // Buyer id column — the Vendor Rating sheet keys sellers via seller_id and
      // (when present) buyers via buyer_id. Detect it so buyer rows are labelled
      // 'buyer' rather than falling through to 'seller'/'combined'.
      var vBuyC = qualityFindCol_(vrh, ['buyer_id','buyerid']);

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

        // — OSV consent (OMP-onboarded sellers only; three-way: verified /
        //   in-progress / not-initiated). The source column carries accepted,
        //   pending and not-yet-started states — classify each so "In Progress"
        //   reflects vendors mid-verification instead of folding them into
        //   not-initiated. A blank cell means verification never started. —
        var vid_ = String(row[vidC] || '').trim();
        var ompInfo = ompMap[vid_];
        if (ompInfo) {
          var osRaw = String(row[oC] || '').trim();
          var osUp  = osRaw.toUpperCase().replace(/\s+/g, '_');
          var osvStatus =
            (osUp === 'CONSENT_ACCEPTED' || osUp === 'YES' || osUp === 'Y' || osUp === 'TRUE'
             || osUp === 'DONE' || osUp === 'COMPLETED' || osUp === 'VERIFIED' || osUp === 'POSITIVE')
              ? 'verified'
            : (osUp === 'CONSENT_PENDING' || osUp === 'PENDING' || osUp === 'IN_PROGRESS'
               || osUp === 'INITIATED' || osUp === 'ONGOING' || osUp === 'PROCESSING'
               || osUp === 'SCHEDULED' || osUp === 'VISIT_SCHEDULED' || osUp === 'PARTIAL')
              ? 'in_progress'
            : 'not_initiated';
          if (osvStatus === 'verified') {
            acc['seller'].o.completed   += 1;
            acc['combined'].o.completed += 1;
          } else if (osvStatus === 'in_progress') {
            acc['seller'].o.pending   += 1;
            acc['combined'].o.pending += 1;
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

        // — Doc completeness: per-document submission list + counts —
        // docList carries every document so the Document Management module can
        // render the full checklist; submitted/total/missingDocs stay for the
        // existing quality summaries. Verification status, upload date and a
        // view URL are surfaced when the source provides them (columns named
        // <doc>_verified / <doc>_date / <doc>_url), else left null.
        var submitted = 0, totalDocs = 0, missingDocNames = [], docList = [];
        docIdx.forEach(function(didx, di) {
          if (didx < 0 || didx >= row.length) return;
          totalDocs++;
          var dv = parseInt(row[didx], 10);
          var isSub = (!isNaN(dv) && dv > 0);
          if (isSub) submitted++;
          else missingDocNames.push(DOC_LABELS[di]);
          var vCol = qualityFindCol_(vrh, [DOC_NAMES[di] + '_verified', DOC_NAMES[di] + '_verification', DOC_NAMES[di] + '_status']);
          var dCol = qualityFindCol_(vrh, [DOC_NAMES[di] + '_date', DOC_NAMES[di] + '_uploaded_date', DOC_NAMES[di] + '_upload_date']);
          var uCol = qualityFindCol_(vrh, [DOC_NAMES[di] + '_url', DOC_NAMES[di] + '_link', DOC_NAMES[di] + '_document_url']);
          docList.push({
            key:        DOC_NAMES[di],
            label:      DOC_LABELS[di],
            submitted:  isSub,
            verified:   vCol >= 0 ? String(row[vCol] || '').trim() : null,
            uploadDate: dCol >= 0 ? String(row[dCol] || '').trim() : null,
            url:        uCol >= 0 ? String(row[uCol] || '').trim() : null
          });
        });
        ts.forEach(function(t) {
          acc[t].d.total += 1;
          if      (totalDocs > 0 && submitted === totalDocs) acc[t].d.complete   += 1;
          else if (submitted > 0)                            acc[t].d.partial    += 1;
          else                                               acc[t].d.incomplete += 1;
        });
        // Enrich with OMP entity info (GSTIN / category / onboarding status) so
        // the repository can show it without a second lookup. Matched by id, then GSTIN.
        var docGst    = vrGstC >= 0 ? String(row[vrGstC] || '').trim() : '';
        var sellerIdV = String(row[vidC] || '').trim();
        var buyerIdV  = vBuyC >= 0 ? String(row[vBuyC] || '').trim() : '';
        // Prefer an explicit id column that is actually populated for this row.
        var docVid    = buyerIdV || sellerIdV;
        // Audience: buyer_id populated → buyer; else the row's resolved audience.
        var docAud    = buyerIdV ? 'buyer' : (sellerIdV ? (rowAud === 'buyer' ? 'buyer' : 'seller') : rowAud);
        var ompD      = ompMap[docVid] || (docGst ? ompGstinMap[docGst] : null);
        if (ompD && ompD.aud) docAud = ompD.aud;
        vendorDocs.push({
          id:               docVid.slice(0, 30),
          name:             (ompD && ompD.name) || String(row[nameC] || '').trim().slice(0, 60),
          gstin:            (ompD && ompD.gstin) || docGst.slice(0, 20),
          category:         ompD ? (ompD.category || '') : '',
          onboardingStatus: ompD ? (ompD.onboardingStatus || '') : '',
          omp:              !!ompD,
          submitted:        submitted,
          total:            totalDocs,
          missingDocs:      missingDocNames,
          docs:             docList,
          aud:              docAud
        });
      });
    }
  } catch (e) { Logger.log('getQualityData vendor-rating: ' + e.message); }

  if (typeof vendorDocs === 'undefined' || !vendorDocs) vendorDocs = [];

  // Buyer documents come from the buyer onboarding sheet (_mb_buyers) — the
  // authoritative source for every Open Marketplace buyer. When it yields buyer
  // records, they REPLACE any buyer rows sourced above (Vendor Rating is
  // seller-centric); if it yields none, existing buyer rows are kept (no
  // regression).
  try {
    var buyerDocs = buildBuyerDocsFromMainSheet_(ompMap, ompGstinMap);
    if (buyerDocs.length) {
      vendorDocs = vendorDocs.filter(function(v) { return v.aud !== 'buyer'; }).concat(buyerDocs);
    }
  } catch (e) { Logger.log('getQualityData buyer-docs (_mb_buyers): ' + e.message); }

  // Supplement any remaining gaps from the dedicated "Doc Completeness" sheet
  // (card 5674). Entities already present (matched by GSTIN, else id) are kept.
  try {
    appendDocsFromCompletenessSheet_(vendorDocs, ompMap, ompGstinMap, sellerIds, buyerIds);
  } catch (e) { Logger.log('getQualityData doc-completeness: ' + e.message); }

  // Rating denominator = OMP-onboarded count per audience (GSTIN-matched)
  acc['seller'].r.total   = Object.keys(ompGstinSellers).length;
  acc['buyer'].r.total    = Object.keys(ompGstinBuyers).length;
  acc['combined'].r.total = Object.keys(ompGstinSellers).length + Object.keys(ompGstinBuyers).length;

  // OSV denominator = OMP-onboarded count. Not-initiated is whatever remains
  // after the verified (completed) and in-progress (pending) vendors, so the
  // three bands always sum to the total.
  ['seller', 'combined'].forEach(function(t) {
    acc[t].o.total        = ompTotal;
    acc[t].o.notInitiated = Math.max(0, ompTotal - acc[t].o.completed - acc[t].o.pending);
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

// Read the "Doc Completeness" sheet (card 5674) and add per-vendor document
// records — importantly, BUYERS, which the seller-centric Vendor Rating sheet
// lacks. Skips entities already present (matched by GSTIN, else id). Handles the
// per-document-flag layout (one 0/1 column per document); if that layout isn't
// present, adds nothing (safe) and debugDocs() will reveal the real schema.
function appendDocsFromCompletenessSheet_(vendorDocs, ompMap, ompGstinMap, sellerIds, buyerIds) {
  var d = qualityReadSheet_('Doc Completeness');
  if (!d.rows.length) return;
  var h = d.headers;
  function fc(c) { return qualityFindCol_(h, c); }
  var sIdC = fc(['seller_id']);
  var bIdC = fc(['buyer_id']);
  var idC  = fc(['vendor_id','id','vendorid','entity_id']);
  var nmC  = fc(['seller_name','buyer_name','vendor_name','name','business_name','company_name']);
  var gstC = fc(['gstin','gst_number','gst_no','gstin_number','gst']);
  var audC = fc(['audience','type','aud','entity_type','seller_buyer','vendor_type']);
  var catC = fc(['business_category','category','vertical_category','cat']);
  var stC  = fc(['onboarding_status','status','onboard_status']);
  var docIdx = DOC_NAMES.map(function(n) { return docColIndex_(h, n); });
  if (docIdx.every(function(i) { return i < 0; })) return;   // no per-document columns → nothing to add

  // Index existing entities so we don't double-add sellers already parsed above.
  var seen = {};
  vendorDocs.forEach(function(v) {
    if (v.gstin) seen['g\x00' + String(v.gstin).toLowerCase()] = true;
    if (v.id)    seen['i\x00' + String(v.id)] = true;
  });

  d.rows.forEach(function(row) {
    var sId = sIdC >= 0 ? String(row[sIdC] || '').trim() : '';
    var bId = bIdC >= 0 ? String(row[bIdC] || '').trim() : '';
    var gId = idC  >= 0 ? String(row[idC]  || '').trim() : '';
    var gst = gstC >= 0 ? String(row[gstC] || '').trim() : '';
    var vid = bId || sId || gId;
    if (!vid && !gst) return;
    if ((gst && seen['g\x00' + gst.toLowerCase()]) || (vid && seen['i\x00' + vid])) return;   // already have it

    // Resolve audience: explicit column → id column → id-set membership → OMP map.
    var aud = '';
    if (audC >= 0) {
      var av = String(row[audC] || '').toLowerCase();
      if (av.indexOf('buy') !== -1) aud = 'buyer';
      else if (av.indexOf('sell') !== -1) aud = 'seller';
    }
    if (!aud) aud = bId ? 'buyer' : (sId ? 'seller' : '');
    if (!aud && vid) aud = buyerIds[vid] ? 'buyer' : (sellerIds[vid] ? 'seller' : '');
    var ompD = ompMap[vid] || (gst ? ompGstinMap[gst] : null);
    if (!aud && ompD) aud = ompD.aud;
    if (!aud) aud = 'combined';

    var submitted = 0, totalDocs = 0, missing = [], docList = [];
    docIdx.forEach(function(di, i) {
      if (di < 0 || di >= row.length) return;
      totalDocs++;
      var dv = parseInt(row[di], 10);
      var isSub = (!isNaN(dv) && dv > 0);
      if (isSub) submitted++; else missing.push(DOC_LABELS[i]);
      var vC = fc([DOC_NAMES[i] + '_verified', DOC_NAMES[i] + '_verification', DOC_NAMES[i] + '_status']);
      var dC = fc([DOC_NAMES[i] + '_date', DOC_NAMES[i] + '_uploaded_date', DOC_NAMES[i] + '_upload_date']);
      var uC = fc([DOC_NAMES[i] + '_url', DOC_NAMES[i] + '_link', DOC_NAMES[i] + '_document_url']);
      docList.push({
        key: DOC_NAMES[i], label: DOC_LABELS[i], submitted: isSub,
        verified:   vC >= 0 ? String(row[vC] || '').trim() : null,
        uploadDate: dC >= 0 ? String(row[dC] || '').trim() : null,
        url:        uC >= 0 ? String(row[uC] || '').trim() : null
      });
    });

    vendorDocs.push({
      id:               vid.slice(0, 30),
      name:             (ompD && ompD.name) || (nmC >= 0 ? String(row[nmC] || '').trim().slice(0, 60) : ''),
      gstin:            (ompD && ompD.gstin) || gst.slice(0, 20),
      category:         (ompD && ompD.category) || (catC >= 0 ? String(row[catC] || '').trim().slice(0, 60) : ''),
      onboardingStatus: (ompD && ompD.onboardingStatus) || (stC >= 0 ? String(row[stC] || '').trim().toUpperCase() : ''),
      omp:              !!ompD,
      submitted:        submitted,
      total:            totalDocs,
      missingDocs:      missing,
      docs:             docList,
      aud:              aud
    });
    if (gst) seen['g\x00' + gst.toLowerCase()] = true;
    if (vid) seen['i\x00' + vid] = true;
  });
}

// Build buyer document records from the buyer onboarding sheet (_mb_buyers) —
// the authoritative source for every Open Marketplace buyer. Detects the
// document columns (canonical names, aliases, then a token-based fuzzy match for
// the three buyer-mandatory docs) and marks a document submitted when its cell
// holds a flag / URL / date. Returns [] when no document columns are present so
// the caller can safely keep whatever it already had.
function buildBuyerDocsFromMainSheet_(ompMap, ompGstinMap) {
  var out   = [];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_mb_buyers');
  if (!sheet || sheet.getLastRow() < 2) return out;
  var d = readSheetObj_(sheet);
  var h = d.headers;
  function fc(c) { return qualityFindCol_(h, c); }
  var idC  = fc(['buyer_id','id','vendor_id','vendorid']);
  var nmC  = fc(['buyer_name','name','vendor_name','business_name','company_name']);
  var gstC = fc(['gstin','gst_number','gst_no','gstin_number','gst']);
  var catC = fc(['business_category','category','vertical_category','cat']);
  var bvC  = fc(['business_vertical','vertical','biz_vertical']);
  var stC  = fc(['onboarding_status','status','onboard_status']);
  var docIdx = DOC_NAMES.map(function(n) { return docColFuzzy_(h, n); });
  if (docIdx.every(function(i) { return i < 0; })) return out;   // no document columns → nothing to build

  d.rows.forEach(function(row) {
    // Scope to Open Marketplace when the sheet carries a vertical column.
    if (bvC >= 0) {
      var bv = String(row[bvC] || '').trim().toLowerCase();
      if (bv && bv.indexOf('open marketplace') < 0 && bv.indexOf('open_marketplace') < 0 && bv !== 'omp') return;
    }
    var id  = idC  >= 0 ? String(row[idC]  || '').trim() : '';
    var gst = gstC >= 0 ? String(row[gstC] || '').trim() : '';
    if (!id && !gst) return;

    var submitted = 0, totalDocs = 0, missing = [], docList = [];
    docIdx.forEach(function(di, i) {
      if (di < 0 || di >= row.length) return;
      totalDocs++;
      var cell  = row[di];
      var isSub = isDocSubmitted_(cell);
      if (isSub) submitted++; else missing.push(DOC_LABELS[i]);
      var url = (typeof cell === 'string' && /^https?:\/\//i.test(cell.trim())) ? cell.trim() : null;
      docList.push({ key: DOC_NAMES[i], label: DOC_LABELS[i], submitted: isSub, verified: null, uploadDate: null, url: url });
    });

    var ompD = ompMap[id] || (gst ? ompGstinMap[gst] : null);
    out.push({
      id:               id.slice(0, 30),
      name:             (nmC >= 0 ? String(row[nmC] || '').trim().slice(0, 60) : '') || (ompD && ompD.name) || '',
      gstin:            gst.slice(0, 20) || (ompD && ompD.gstin) || '',
      category:         (catC >= 0 ? String(row[catC] || '').trim().slice(0, 60) : '') || (ompD ? ompD.category : ''),
      onboardingStatus: (stC  >= 0 ? String(row[stC]  || '').trim().toUpperCase() : '') || (ompD ? ompD.onboardingStatus : ''),
      omp:              true,
      submitted:        submitted,
      total:            totalDocs,
      missingDocs:      missing,
      docs:             docList,
      aud:              'buyer'
    });
  });
  return out;
}

// Manual diagnostic — run from the Apps Script editor to see where document
// data lives and how it splits by audience. Logs the headers + counts of the
// Vendor Rating, Doc Completeness and _mb_buyers sheets and the final
// per-audience vendorDocs breakdown, so buyer-document sourcing can be verified.
function debugDocs() {
  Logger.log('════ DOC DEBUG ════');
  ['Vendor Rating', 'Doc Completeness', '_mb_buyers'].forEach(function(nm) {
    var s = qualityReadSheet_(nm);
    Logger.log('— "' + nm + '": ' + s.rows.length + ' rows');
    if (s.headers.length) {
      Logger.log('   headers: ' + s.headers.join(', '));
      var found = DOC_NAMES.filter(function(n) { return s.headers.indexOf(n) !== -1; });
      Logger.log('   doc columns present (by name): ' + (found.length ? found.join(', ') : '(none)'));
      // Show how the three buyer-mandatory docs resolve (alias + fuzzy).
      ['gst_certificate', 'entity_pan', 'pwm_certificate'].forEach(function(k) {
        var ix = docColFuzzy_(s.headers, k);
        Logger.log('   mandatory "' + k + '" → ' + (ix >= 0 ? 'col "' + s.headers[ix] + '"' : '(not found)'));
      });
      Logger.log('   id cols → seller_id:' + (s.headers.indexOf('seller_id') !== -1)
        + ' buyer_id:' + (s.headers.indexOf('buyer_id') !== -1)
        + ' audience/type:' + (s.headers.indexOf('audience') !== -1 || s.headers.indexOf('type') !== -1));
    }
  });
  try {
    var q = JSON.parse(getQualityData());
    var vd = q.vendorDocs || [];
    var byAud = { seller: 0, buyer: 0, combined: 0, other: 0 };
    vd.forEach(function(v) { byAud[v.aud] = (byAud[v.aud] || 0) + 1; });
    Logger.log('vendorDocs total: ' + vd.length + ' → seller:' + (byAud.seller || 0)
      + ' buyer:' + (byAud.buyer || 0) + ' combined:' + (byAud.combined || 0));
    var b = vd.filter(function(v) { return v.aud === 'buyer'; }).slice(0, 5);
    b.forEach(function(v) { Logger.log('   buyer sample: ' + v.name + ' | ' + v.gstin + ' | ' + v.submitted + '/' + v.total); });
    if (!(byAud.buyer)) Logger.log('   ⚠ No buyer document rows — check the sheets/columns logged above.');
  } catch (e) { Logger.log('getQualityData failed: ' + e.message); }
  Logger.log('════ END ════');
}
