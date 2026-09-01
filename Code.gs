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

  // Workbook the Metabase tabs are synced into. TAT reads _mb_detail from here by
  // id rather than from the active spreadsheet, so the source is the same whichever
  // document the script is opened from. Falls back to the active spreadsheet.
  META_SHEET_ID: '10RJ1D1GXh-f_7a5M3YMAEt8jDQ7X6jQm2-krTNOBts8',

  // External workbook holding Vendor Rating + OSV Status. Score / OSV columns are now
  // resolved by HEADER NAME in getQualityData (positional VS_COLS kept only as a last
  // fallback). Run debugVendorScoreSheet() to dump the live headers and confirm the
  // rating / OSV / GSTIN columns are detected.
  VENDOR_SCORE_SHEET_ID: '1UMtuarqR9wFI74VM4JWC3GXSF9C9YpkkySeFJgq8rQc',   // "NBFC TRACKER"
  VENDOR_SCORE_TAB:      '',   // blank → first tab in the workbook
};

// Fixed column positions in the Vendor Score workbook (0-based indices).
var VS_COLS = {
  denominator: 7,    // Column H — Vendor Scores denominator
  osvStatus:   13,   // Column N — OSV Status
};

// Buyer TAT is read positionally from the buyer feed (_mb_buyers): column G is the
// start and column H the end, as specified. Addressed by position rather than header
// name because that is how they were given. Run debugBuyerTatCols() to confirm what
// actually sits at those positions and the span they produce.
var BUYER_TAT_COLS = {
  start: 6,   // Column G
  end:   7,   // Column H
};

// Candidate column names in the 5292 detail tab (headers are normalized to
// lower_snake_case by readSheetObj_). First populated match wins per record.
var TAT_COLS = {
  // TAT start, preferred: the moment the case entered review, per card 5292.
  inReview:  ['in_review_date','in_review_at','in_review','under_review_date','under_review_at',
              'review_started_date','review_started_at','review_initiated_date','review_initiated_at',
              'review_date','review_at','submitted_date','submitted_at','submission_date',
              'onboarding_in_review_date','onboarding_submitted_date','sent_for_review_date',
              'sent_for_review_at','moved_to_review_date'],
  // TAT start, fallback: latest Level 1 date.
  level1:    ['latest_level_1','latest_level_1_date','level_1_date','level_1_at','level_1_on',
              'level_1','level1_date','level1','l1_date','l1','level_1_approval_date',
              'level_1_review_date','level_1_updated_date','level_1_completed_date'],
  // A case that was rejected and sent back restarts the clock: time spent waiting
  // for the vendor to correct and resubmit is not review turnaround.
  resubmitted: ['resubmitted_date','resubmitted_at','resubmission_date','resubmission_at',
                'reapplied_date','reapplied_at','resubmit_date','last_resubmitted_date',
                're_review_date','sent_back_resubmitted_date'],
  rejected:    ['rejected_date','rejected_at','rejection_date','rejection_at',
                'declined_date','declined_at','sent_back_date','returned_date'],
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
  { key: 'OMP',           name: 'Marketplace',         code: 'OMP', sub: '' },
  { key: 'EPR',           name: 'EPR',                 code: 'EPR', sub: '' },
  { key: 'Marketplace',   name: 'Managed Marketplace', code: 'MKT', sub: '' },
  { key: 'InfraBusiness', name: 'Infra Business',      code: 'INF', sub: '' },
  { key: 'AFR',           name: 'AFR',                 code: 'AFR', sub: '' },
  { key: 'Recommerce',    name: 'Re-Commerce',         code: 'REC', sub: '' },
  { key: 'DRS',           name: 'DRS',                 code: 'DRS', sub: '' },
  { key: 'Others',        name: 'Others',              code: 'OTH', sub: '' },
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
    // Non-infra Marketplace categories (Paper, M3, M4, Tyre Oil, etc.) belong
    // directly in Managed Marketplace — no indirection through Others.
    return 'Marketplace';
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
    var cacheKey  = 'dash_v36_' + audience + '_' + periodKey;
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
      if (v.length <= 100000) batchCache['vrows_v17_' + audience + '_' + vc.key + '_' + periodKey] = v;
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
// Fast path: compose from individual dash_v36_ cache entries when both are warm
// (they are pre-warmed by syncAllOnboarding for every period). Only falls through
// to the slow double-read when both individual caches are cold.
function getCombinedDashboard(filtersJson) {
  try {
    var f = filtersJson ? JSON.parse(filtersJson) : {};
    var periodKey = JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cacheKey  = 'dash_v36_cmb_' + periodKey;
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    // Try to compose from pre-warmed individual caches (zero extra reads).
    var sIndKey = 'dash_v36_seller_' + periodKey;
    var bIndKey = 'dash_v36_buyer_'  + periodKey;
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

    var cacheKey = 'vrows_v17_' + audience + '_' + vertKey + '_'
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
      var cacheKey = 'vrows_v17_' + aud + '_OMP_' + periodKey;
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

// Locate a date column in the detail sheet: the named candidates first, then a
// scan for any header matching `fuzzy` whose values actually parse as dates.
// Returns { i, header } or null. The scan exists because the detail card's column
// names vary, and a missed column silently costs every record its TAT.
function _detailDateCol_(headers, rows, exact, fuzzy) {
  var idx = buildIndex(headers);
  for (var e = 0; e < exact.length; e++) {
    if (idx[exact[e]] !== undefined) return { i: idx[exact[e]], header: exact[e] };
  }
  var best = null;
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (!h || !fuzzy.test(h)) continue;
    if (/_by$|user|name|status|remark|comment|reason|count|id$/.test(h)) continue;
    var hits = 0, seen = 0;
    for (var r = 0; r < rows.length && seen < 200; r++) {
      var v = rows[r][c];
      if (v === '' || v === null || v === undefined) continue;
      seen++;
      if (parseDate(v)) hits++;
    }
    if (hits > 0 && (!best || hits > best.hits)) best = { i: c, header: h, hits: hits };
  }
  return best;
}

// Per-vendor TAT dates read from the card 5292 detail tab — the dedicated
// onboarding-stage source. Keyed by id, and by name where that name maps to a
// single id. Holds the In Review date (preferred start), the Level 1 date
// (fallback start) and the detail sheet's own completed date.
// Open the _mb_detail tab from the configured meta workbook, falling back to the
// active spreadsheet. Returns the Sheet or null.
function openDetailSheet_() {
  var name = CONFIG.TAT_DETAIL_SHEET;
  if (CONFIG.META_SHEET_ID) {
    try {
      var s = SpreadsheetApp.openById(CONFIG.META_SHEET_ID).getSheetByName(name);
      if (s && s.getLastRow() > 1) return s;
    } catch (e) { Logger.log('meta workbook open failed: ' + e.message); }
  }
  try { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); } catch (e) { return null; }
}

function getLevel1Lookup_() {
  if (_level1LookupCache) return _level1LookupCache;
  var lk = { byId: {}, byName: {}, cols: {}, rows: 0,
             audCounts: { seller: 0, buyer: 0, unknown: 0 }, collisions: 0, loaded: false };
  try {
    var sheet = openDetailSheet_();
    if (sheet && sheet.getLastRow() > 1) {
      var d   = readSheetObj_(sheet);
      var idx = buildIndex(d.headers);
      lk.rows = d.rows.length;

      var revC = _detailDateCol_(d.headers, d.rows, TAT_COLS.inReview,    /review|submit/);
      var l1C  = _detailDateCol_(d.headers, d.rows, TAT_COLS.level1,      /level_?1|(^|_)l1(_|$)/);
      var endC = _detailDateCol_(d.headers, d.rows, TAT_COLS.completed,   /complete|onboard/);
      var resC = _detailDateCol_(d.headers, d.rows, TAT_COLS.resubmitted, /resubmit|resubmiss|reappl|re_review/);
      var rejC = _detailDateCol_(d.headers, d.rows, TAT_COLS.rejected,    /reject|declin|sent_back|returned/);
      lk.cols = { review: revC && revC.header, level1: l1C && l1C.header,
                  completed: endC && endC.header, resubmitted: resC && resC.header,
                  rejected: rejC && rejC.header };

      // This tab holds BOTH sellers and buyers, so every key is namespaced by
      // audience. A seller_id and a buyer_id can be the same number for different
      // entities; keying on the bare id would hand one audience the other's dates.
      var sIdC = idx['seller_id'] !== undefined ? idx['seller_id'] : -1;
      var bIdC = idx['buyer_id']  !== undefined ? idx['buyer_id']  : -1;
      var audC = -1;
      ['audience','entity_type','type','vendor_type_group','party_type','user_type'].forEach(function(k) {
        if (audC < 0 && idx[k] !== undefined) audC = idx[k];
      });
      // Row audience: an explicit column wins; otherwise whichever id column is filled.
      function rowAud_(row) {
        if (audC >= 0) {
          var a = String(row[audC] || '').trim().toLowerCase();
          if (a.indexOf('sell') === 0 || a === 's' || a.indexOf('vendor') === 0) return 'seller';
          if (a.indexOf('buy')  === 0 || a === 'b' || a.indexOf('customer') === 0) return 'buyer';
        }
        var hasS = sIdC >= 0 && String(row[sIdC] || '').trim() !== '';
        var hasB = bIdC >= 0 && String(row[bIdC] || '').trim() !== '';
        if (hasS && !hasB) return 'seller';
        if (hasB && !hasS) return 'buyer';
        return 'unknown';
      }
      // Id for a row, preferring the column matching its audience.
      function rowId_(row, aud) {
        if (aud === 'seller' && sIdC >= 0) {
          var s = String(row[sIdC] || '').replace(/,/g, '').trim(); if (s) return s;
        }
        if (aud === 'buyer' && bIdC >= 0) {
          var b = String(row[bIdC] || '').replace(/,/g, '').trim(); if (b) return b;
        }
        return String(firstVal_(row, idx, TAT_COLS.id) || '').replace(/,/g, '').trim();
      }

      var nameKeys = {};   // "aud|name" → {id, ambiguous}: a name used by >1 id is unsafe
      // Keep the LATEST date seen per vendor: the detail card can emit one row per
      // stage transition, and the most recent pass is the one that led to onboarding.
      function fold_(map, key, field, dt) {
        if (!dt) return;
        var rec = map[key] || (map[key] = { review: null, level1: null, completed: null,
                                            resubmitted: null, rejected: null });
        if (!rec[field] || dt > rec[field]) rec[field] = dt;
      }
      function foldAll_(map, key, dates) {
        fold_(map, key, 'review', dates.rev);   fold_(map, key, 'level1', dates.l1);
        fold_(map, key, 'completed', dates.end); fold_(map, key, 'resubmitted', dates.res);
        fold_(map, key, 'rejected', dates.rej);
      }
      d.rows.forEach(function(row) {
        var dates = {
          rev: revC ? parseDate(row[revC.i]) : null,
          l1:  l1C  ? parseDate(row[l1C.i])  : null,
          end: endC ? parseDate(row[endC.i]) : null,
          res: resC ? parseDate(row[resC.i]) : null,
          rej: rejC ? parseDate(row[rejC.i]) : null
        };
        if (!dates.rev && !dates.l1 && !dates.end && !dates.res && !dates.rej) return;
        var aud  = rowAud_(row);
        lk.audCounts[aud]++;
        var id   = rowId_(row, aud);
        var name = String(firstVal_(row, idx, TAT_COLS.name) || '').trim().toLowerCase();
        // Strictly namespaced. A row whose audience IS known is reachable only under
        // that audience — no shared bucket, because an id appearing for sellers only
        // would otherwise still answer a buyer lookup and hand over the wrong dates.
        // Rows whose audience cannot be determined go to 'unknown', which both
        // audiences may read since there is nothing to distinguish them by.
        if (id) foldAll_(lk.byId, aud + '|' + id, dates);
        if (name) {
          foldAll_(lk.byName, aud + '|' + name, dates);
          var nk = nameKeys[aud + '|' + name];
          if (!nk) nameKeys[aud + '|' + name] = { id: id, ambiguous: false };
          else if (nk.id !== id) nk.ambiguous = true;
        }
      });
      // Count ids present for BOTH audiences — proof the namespacing is earning its
      // keep, since a bare-id lookup would have merged these.
      Object.keys(lk.byId).forEach(function(k) {
        if (k.indexOf('seller|') === 0 && lk.byId['buyer|' + k.slice(7)]) lk.collisions++;
      });
      // Names shared by more than one id are unsafe to match on.
      Object.keys(nameKeys).forEach(function(k) { if (nameKeys[k].ambiguous) delete lk.byName[k]; });
      lk.loaded = Object.keys(lk.byId).length + Object.keys(lk.byName).length > 0;
    }
  } catch (e) { Logger.log('TAT detail lookup build failed: ' + e.message); }
  _level1LookupCache = lk;
  return lk;
}

// Full TAT detail record for a vendor, scoped to its audience so a seller and a
// buyer sharing an id can never resolve to each other. Order: this audience's id,
// then an id whose audience was indeterminate, then this audience's name.
function lookupTatDetail_(id, name, audience) {
  var lk  = getLevel1Lookup_();
  var aud = audience === 'buyer' ? 'buyer' : audience === 'seller' ? 'seller' : null;
  if (id) {
    var k = String(id).replace(/,/g, '').trim();
    if (k) {
      if (aud && lk.byId[aud + '|' + k]) return lk.byId[aud + '|' + k];
      // Only rows whose audience could not be determined are shared between the two.
      if (lk.byId['unknown|' + k])       return lk.byId['unknown|' + k];
    }
  }
  if (name) {
    var n = String(name).trim().toLowerCase();
    if (n) {
      if (aud && lk.byName[aud + '|' + n]) return lk.byName[aud + '|' + n];
      if (lk.byName['unknown|' + n])       return lk.byName['unknown|' + n];
    }
  }
  return null;
}

// Latest Level 1 date for a record (Date), matched by id first then name, or null.
function lookupLevel1_(id, name, audience) {
  var rec = lookupTatDetail_(id, name, audience);
  return rec ? rec.level1 : null;
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
      var basis = { review: 0, level1: 0, created: 0, none: 0 }, completed = 0;
      rows.forEach(function(r) {
        if (lookupLevel1_(r.id, r.name, aud)) l1matched++;
        if (r.status === 'COMPLETED') { completed++; basis[r.tatBasis || 'none']++; }
        if (r.onbTAT !== null) { sumTat += r.onbTAT; n++; }
        if (r.status === 'COMPLETED' && r.createdDate && r.onboardedDate) {
          var b = dateDiffDays(r.createdDate, r.onboardedDate);
          if (b !== null && b >= 0 && b <= TAT_MAX_DAYS) { sumBase += b; nBase++; }
        }
      });
      Logger.log(aud + ': ' + rows.length + ' records · Level 1 matched ' + l1matched
        + ' · Avg TAT (InReview→Onb) ' + (n ? Math.round(sumTat / n) + 'd over ' + n : '—')
        + ' · baseline (Created→Onb) ' + (nBase ? Math.round(sumBase / nBase) + 'd over ' + nBase : '—'));
      Logger.log('  TAT basis over ' + completed + ' completed → review: ' + basis.review
        + ' · level1: ' + basis.level1 + ' · created: ' + basis.created
        + ' · NO START (reports —): ' + basis.none);
      if (basis.review === 0 && (basis.level1 > 0 || basis.created > 0)) {
        Logger.log('  ⚠ ' + aud + ' resolved NO In Review date — TAT is measured from '
          + (basis.level1 > 0 ? 'Level 1' : 'Created') + ' instead and the panel says so.'
          + ' Check the header list below for the real in-review column and add it to'
          + ' the reviewDate candidates in normalizeRows().');
      }
      // Which columns even look like a review timestamp, to make the gap actionable.
      try {
        var _rawH = readData(aud).headers || [];
        var _cand = _rawH.filter(function(h) {
          return /review|submit|level_?1|_l1|in_review/.test(String(h).toLowerCase());
        });
        Logger.log('  headers mentioning review/submit/level1: '
          + (_cand.length ? _cand.join(', ') : '(none — the feed has no such column)'));
      } catch (e) {}
      // Per-category coverage — a category whose average looks wrong is usually one
      // where only a few onboarded records carry a review date.
      var byCat = {};
      rows.forEach(function(r) {
        if (r.status !== 'COMPLETED') return;
        var c = r.category || 'Others';
        if (!byCat[c]) byCat[c] = { onb: 0, withTat: 0, sum: 0 };
        byCat[c].onb++;
        if (r.onbTAT !== null) { byCat[c].withTat++; byCat[c].sum += r.onbTAT; }
      });
      Object.keys(byCat).sort().forEach(function(c) {
        var b = byCat[c];
        var pct = b.onb ? Math.round(b.withTat / b.onb * 100) : 0;
        Logger.log('    ' + c + ': ' + b.withTat + '/' + b.onb + ' onboarded have a TAT (' + pct + '%)'
          + ' · avg ' + (b.withTat ? Math.round(b.sum / b.withTat) + 'd' : '—')
          + (b.withTat > 0 && pct < 40 ? '   ⚠ thin sample — average is unreliable' : ''));
      });
    } catch (e) { Logger.log(aud + ' match check failed: ' + e.message); }
  });
  Logger.log('════ END ════');
}

// ─────────────────────────────────────────────────────────────
// TAT assignment — ONE start date for the whole feed.
//
// TAT is meant to measure In Review → Onboarded. Where a feed carries no review
// timestamp at all, the only remaining signal is the Level 1 date from card 5292.
// Choosing between them PER RECORD is what previously inflated one category
// against another: some records measured the short review span, their neighbours
// the much longer level1 span, and the two were averaged together.
//
// So the basis is decided once, for the entire feed, by whichever start resolves
// for more completed records — and then applied uniformly. Every figure in a feed
// is therefore measured the same way and categories stay comparable, while a feed
// lacking review dates still reports a real turnaround instead of a column of "—".
// The chosen basis travels with the rows so the UI can label what it is showing.
// ─────────────────────────────────────────────────────────────
// Calendar-day number, ignoring time-of-day. Sheets hands back Date objects with a
// live time component for datetime cells, so comparing them raw rejected valid
// records: a case reviewed at 14:30 and onboarded the same day (stored at midnight)
// failed "review <= onboarded" despite being a legitimate 0-day TAT. dateDiffDays
// already works in calendar days — the window guard now agrees with it.
function _dayNum_(d) {
  return d ? Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000 : null;
}
// Last TAT resolution per audience — basis, per-candidate coverage and the reason
// tally. Populated by _assignTat_ and read by the diagnostics.
var _lastTatStats = {};
function _assignTat_(rows, audience) {
  // Measured against _tatEnd (the detail tab's completion date where available),
  // not onboardedDate, which may be a drifting last-touched timestamp.
  // Buyer column H is the end when present, ahead of every other candidate.
  function endOf(r)  { return r._be || r._tatEnd || r.onboardedDate; }
  // A start is usable if it lands on or before the end. For inferred starts we also
  // require it to fall on or after the created date, as a sanity check against a
  // stray column. That check is NOT applied to the explicitly specified buyer
  // columns: G/H were given as the source, and requiring G >= onboarding_created_date
  // silently rejected most buyers whose G sits earlier in the journey.
  function inWin(r, d, basis) {
    var dn = _dayNum_(d), on = _dayNum_(endOf(r));
    if (dn === null || on === null || dn > on) return false;
    if (basis === 'fixed') return true;
    var cn = _dayNum_(r.createdDate);
    return cn === null || dn >= cn;
  }
  // The 'fixed' basis is the buyer feed's columns G→H, addressed by position.
  function startOf(r, basis) {
    return basis === 'fixed'  ? r._bs
         : basis === 'review' ? r.reviewDate
         : basis === 'level1' ? r._l1
         : r.createdDate;
  }
  var done = rows.filter(function(r) { return r.status === 'COMPLETED' && endOf(r); });
  var n = { review: 0, level1: 0, created: 0, fixed: 0 };
  done.forEach(function(r) {
    if (inWin(r, r.reviewDate,  'review'))  n.review++;
    if (inWin(r, r._l1,         'level1'))  n.level1++;
    if (inWin(r, r.createdDate, 'created')) n.created++;
    if (inWin(r, r._bs,         'fixed'))   n.fixed++;
  });

  // In Review → Onboarded is the metric asked for, so review wins whenever it covers
  // a real share of the feed. Below that it would describe too few records to stand
  // as the vertical's turnaround, and the next-best start with wider coverage is used
  // instead — Level 1, then Created. Whatever is chosen is applied to EVERY record in
  // the feed, so all figures share one basis and categories stay comparable; the mixed
  // per-record fallback is what previously inflated one category against another.
  var basis = null;
  if (audience === 'buyer') {
    // Buyers are measured from the buyer feed's own columns G → H. Created → Onboarded
    // remains only as a fallback for feeds where those positions hold no usable dates.
    basis = n.fixed > 0 ? 'fixed' : (n.created > 0 ? 'created' : null);
  }
  else if (done.length && n.review >= done.length * 0.5) basis = 'review';
  else if (n.level1  > n.review && n.level1  > 0)        basis = 'level1';
  else if (n.review  > 0)                                basis = 'review';
  else if (n.created > 0)                                basis = 'created';

  // Single choke point for what may carry a TAT. Both rules are enforced here so no
  // consumer can diverge: only ONBOARDED (COMPLETED) vendors are measured, and a span
  // outside 0..TAT_MAX_DAYS is never stored. Anything else keeps onbTAT null, so
  // downstream code needs only a null check and the records table, the averages and
  // the category splits can no longer disagree about which records count.
  // Tally every reason an onboarded record ends up without a TAT, so a shortfall
  // like "19 of 78 buyers" can be attributed instead of guessed at.
  var tally = { onboarded: 0, ok: 0, noStart: 0, noEnd: 0, startAfterEnd: 0, tooLong: 0, noBasis: 0 };
  rows.forEach(function(r) {
    if (r.status === 'COMPLETED') {
      tally.onboarded++;
      var end = endOf(r), start = basis ? startOf(r, basis) : null;
      if (!basis)                        tally.noBasis++;
      else if (!end)                     tally.noEnd++;
      else if (!start)                   tally.noStart++;
      else if (!inWin(r, start, basis))  tally.startAfterEnd++;
      else {
        var t = dateDiffDays(start, end);
        if (t === null || t < 0)         tally.startAfterEnd++;
        else if (t > TAT_MAX_DAYS)       tally.tooLong++;
        else { r.onbTAT = t; r.tatBasis = basis; tally.ok++; }
      }
    }
    delete r._l1;       // internal candidates — never reach a payload or cache
    delete r._tatEnd; delete r._bs; delete r._be;
  });
  _lastTatStats[audience || 'seller'] = { basis: basis, coverage: n, tally: tally };
  return basis;
}

// ─────────────────────────────────────────────────────────────
// Run from the Apps Script editor. For each audience this lists EVERY column that
// parses as a date on onboarded records, how many it covers, and the TAT it would
// produce if used as the start. That turns "which column is the In Review date?"
// from a guess into a decision: pick the column whose name matches the milestone
// and whose median TAT is plausible, then add it to the reviewDate candidates.
// ─────────────────────────────────────────────────────────────
function debugTATColumns() {
  ['seller', 'buyer'].forEach(function(aud) {
    Logger.log('\n════════ ' + aud.toUpperCase() + ' ════════');
    var raw, cfg = AUDIENCE_CFG[aud];
    try { raw = readData(aud); } catch (e) { Logger.log('  ✗ readData failed: ' + e.message); return; }
    var idx = buildIndex(raw.headers);
    var onbCol = cfg.onbCol, stCol = 'onboarding_status';
    if (idx[onbCol] === undefined) { Logger.log('  ✗ onboarded column "' + onbCol + '" not present'); return; }

    // Onboarded rows only — TAT is defined for completed records.
    var done = raw.rows.filter(function(r) {
      return normStatus(gv(r, idx, stCol)) === 'COMPLETED' && parseDate(r[idx[onbCol]]);
    });
    Logger.log('  onboarded rows: ' + done.length + '   (end date column: "' + onbCol + '")');
    if (!done.length) return;

    Logger.log('  ── every column that parses as a date, as a candidate START ──');
    var results = [];
    raw.headers.forEach(function(h, ci) {
      if (!h || ci === idx[onbCol]) return;
      var tats = [], parsed = 0;
      done.forEach(function(r) {
        var d = parseDate(r[ci]); if (!d) return;
        parsed++;
        var end = parseDate(r[idx[onbCol]]);
        var t = dateDiffDays(d, end);
        if (t !== null && t >= 0 && t <= TAT_MAX_DAYS) tats.push(t);
      });
      if (!parsed) return;
      tats.sort(function(a, b) { return a - b; });
      var med = tats.length ? tats[Math.floor(tats.length / 2)] : null;
      var mean = tats.length ? Math.round(tats.reduce(function(a, b) { return a + b; }, 0) / tats.length) : null;
      results.push({ h: h, col: colLetter_(ci), parsed: parsed, usable: tats.length,
                     med: med, mean: mean, min: tats[0], max: tats[tats.length - 1] });
    });
    results.sort(function(a, b) { return b.usable - a.usable; });
    results.forEach(function(r) {
      Logger.log('   ' + r.col.padEnd(4) + (r.h + '                              ').slice(0, 30)
        + ' parses ' + String(r.parsed).padStart(5) + '/' + done.length
        + ' · usable ' + String(r.usable).padStart(5)
        + ' · median ' + (r.med != null ? String(r.med) + 'd' : '—')
        + ' · mean ' + (r.mean != null ? String(r.mean) + 'd' : '—')
        + ' · range ' + (r.usable ? r.min + '–' + r.max + 'd' : '—'));
    });
    Logger.log('  ── what the dashboard currently uses ──');
    var rows = normalizeRows(raw, cfg);
    var used = {}, nT = 0, sT = 0;
    rows.forEach(function(r) {
      if (r.status !== 'COMPLETED') return;
      used[r.tatBasis || 'none'] = (used[r.tatBasis || 'none'] || 0) + 1;
      if (r.onbTAT !== null) { nT++; sT += r.onbTAT; }
    });
    Logger.log('   basis split: ' + JSON.stringify(used)
      + '   ·   avg TAT ' + (nT ? Math.round(sT / nT) + 'd over ' + nT : '—'));
  });
  Logger.log('\nPick the column matching your In Review milestone and add its header to the'
    + ' reviewDate candidates in normalizeRows(). Buyers are pinned to Created by design.');
}

// ─────────────────────────────────────────────────────────────
// Run from the Apps Script editor. Scans the card 5292 detail tab that TAT is now
// sourced from: which columns were auto-detected as In Review / Level 1 / Completed,
// every other column that parses as a date (with the TAT it would produce), and how
// many live seller / buyer records actually match a detail row.
// ─────────────────────────────────────────────────────────────
function debugTatDetail() {
  var name = CONFIG.TAT_DETAIL_SHEET;
  var sheet = openDetailSheet_();
  Logger.log('════ TAT DETAIL TAB "' + name + '" (Metabase card '
    + (CONFIG.MB_CARDS_DETAIL || 5292) + ') ════');
  Logger.log('meta workbook: ' + (CONFIG.META_SHEET_ID || '(active spreadsheet)'));
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('✗ tab missing or empty — run syncDetail() in Metabase.gs first.');
    return;
  }
  var d = readSheetObj_(sheet);
  Logger.log('rows: ' + d.rows.length + '   columns: ' + d.headers.length);
  Logger.log('\n── all headers ──');
  d.headers.forEach(function(h, i) { Logger.log('  ' + colLetter_(i) + '  ' + (h || '(blank)')); });

  _level1LookupCache = null;                       // force a fresh scan
  var lk = getLevel1Lookup_();
  Logger.log('\n── auto-detected TAT columns ──');
  Logger.log('  In Review   : ' + (lk.cols.review      || '✗ NOT FOUND'));
  Logger.log('  Level 1     : ' + (lk.cols.level1      || '✗ not found'));
  Logger.log('  Resubmitted : ' + (lk.cols.resubmitted || '✗ not found — rejected cases will use the latest post-rejection review date'));
  Logger.log('  Rejected    : ' + (lk.cols.rejected    || '✗ not found — cannot tell a stale pre-rejection review from a valid one'));
  Logger.log('  Completed   : ' + (lk.cols.completed   || '✗ not found — TAT falls back to the feed onboarded date'));
  Logger.log('  vendors keyed by id: ' + Object.keys(lk.byId).length
    + ' · by name: ' + Object.keys(lk.byName).length);
  Logger.log('  detail rows by audience: seller ' + lk.audCounts.seller
    + ' · buyer ' + lk.audCounts.buyer + ' · indeterminate ' + lk.audCounts.unknown);
  if (lk.audCounts.buyer === 0) {
    Logger.log('  ⚠ no BUYER rows identified in this tab. If buyers are present, their id'
      + ' column is not named buyer_id and there is no audience column — buyer TAT will'
      + ' fall back to the feed dates.');
  }
  if (lk.collisions) {
    Logger.log('  ' + lk.collisions + ' id(s) appear for BOTH audiences and were dropped from the'
      + ' shared bucket; those vendors resolve only via their own audience key.');
  }

  // Every date-like column and the span it would give against the detail's own end.
  Logger.log('\n── every column that parses as a date ──');
  var endI = lk.cols.completed ? buildIndex(d.headers)[lk.cols.completed] : -1;
  d.headers.forEach(function(h, ci) {
    if (!h) return;
    var parsed = 0, tats = [];
    d.rows.forEach(function(r) {
      var dt = parseDate(r[ci]); if (!dt) return;
      parsed++;
      if (endI >= 0) {
        var e = parseDate(r[endI]);
        var t = e ? dateDiffDays(dt, e) : null;
        if (t !== null && t >= 0 && t <= TAT_MAX_DAYS) tats.push(t);
      }
    });
    if (!parsed) return;
    tats.sort(function(a, b) { return a - b; });
    Logger.log('  ' + colLetter_(ci).padEnd(4) + (h + '                              ').slice(0, 30)
      + ' dates ' + String(parsed).padStart(5) + '/' + d.rows.length
      + (tats.length ? '  · median span ' + tats[Math.floor(tats.length / 2)] + 'd'
                       + ' · range ' + tats[0] + '–' + tats[tats.length - 1] + 'd' : ''));
  });

  // How much of the live data actually joins to this tab.
  Logger.log('\n── match rate against live records ──');
  ['seller', 'buyer'].forEach(function(aud) {
    try {
      var rows = normalizeRows(readData(aud), AUDIENCE_CFG[aud]);
      var done = rows.filter(function(r) { return r.status === 'COMPLETED'; });
      var matched = 0, withRev = 0, withL1 = 0, withTat = 0, sum = 0;
      done.forEach(function(r) {
        var rec = lookupTatDetail_(r.id, r.name, aud);
        if (rec) { matched++; if (rec.review) withRev++; if (rec.level1) withL1++; }
        if (r.onbTAT !== null) { withTat++; sum += r.onbTAT; }
      });
      Logger.log('  ' + aud + ': ' + done.length + ' onboarded · matched a detail row ' + matched
        + ' · of those In Review ' + withRev + ', Level 1 ' + withL1
        + '  →  TAT on ' + withTat + ' records, avg ' + (withTat ? Math.round(sum / withTat) + 'd' : '—'));
    } catch (e) { Logger.log('  ' + aud + ': ' + e.message); }
  });
  Logger.log('\nIf "In Review" reads NOT FOUND, no column in this tab both mentions'
    + ' review/submit and parses as a date — add the real header to TAT_COLS.inReview.');
}

// Index of a genuine completion/approval timestamp in a feed, or -1.
// Deliberately refuses "updated"/"modified"/"last_" columns: the buyer card has no
// onboarded_date and falls back to onboarding_updated_date, which moves whenever a
// record is touched and stretched buyer turnaround well past the real span.
// Shared by normalizeRows and the diagnostics so both agree on the column in use.
function tatEndColIndex_(raw) {
  var idx = buildIndex(raw.headers);
  var exact = ['onboarded_date','onboarding_completed_date','onboarding_completion_date',
               'completed_date','completion_date','onboarded_at','completed_at',
               'onboarding_approved_date','approved_date','approval_date','approved_at',
               'activation_date','activated_date','go_live_date','kyc_approved_date'];
  for (var e = 0; e < exact.length; e++) {
    if (idx[exact[e]] !== undefined) return idx[exact[e]];
  }
  var best = -1, bestHits = 0;
  for (var c = 0; c < raw.headers.length; c++) {
    var h = raw.headers[c];
    if (!h) continue;
    if (!/onboard|complet|approv|activat|go_live/.test(h)) continue;
    if (/updated|modified|last_|create|draft|review|reject|status|_by$|user|name|count|id$/.test(h)) continue;
    var hits = 0, seen = 0;
    for (var r = 0; r < raw.rows.length && seen < 200; r++) {
      var v = raw.rows[r][c];
      if (v === '' || v === null || v === undefined) continue;
      seen++;
      if (parseDate(v)) hits++;
    }
    if (hits > bestHits) { best = c; bestHits = hits; }
  }
  return best;
}

// ─────────────────────────────────────────────────────────────
// Run from the Apps Script editor when a TAT figure looks wrong for one audience.
// Prints, per audience, the start and end columns actually in use and then a grid
// of every plausible start against every plausible end with the median span each
// pair produces — so the correct pair can be read off directly rather than guessed.
// ─────────────────────────────────────────────────────────────
function debugTatPairs() {
  ['seller', 'buyer'].forEach(function(aud) {
    var cfg = AUDIENCE_CFG[aud], raw;
    Logger.log('\n════════ ' + aud.toUpperCase() + ' ════════');
    try { raw = readData(aud); } catch (e) { Logger.log('  ✗ ' + e.message); return; }
    var idx = buildIndex(raw.headers);
    var endC = tatEndColIndex_(raw);
    Logger.log('  counts use cfg.onbCol = "' + cfg.onbCol + '"');
    Logger.log('  TAT end column in use  = '
      + (endC >= 0 ? '"' + raw.headers[endC] + '" (' + colLetter_(endC) + ')'
                   : 'none found → falls back to "' + cfg.onbCol + '"'));

    var done = raw.rows.filter(function(r) {
      return normStatus(gv(r, idx, 'onboarding_status')) === 'COMPLETED';
    });
    Logger.log('  onboarded rows: ' + done.length);
    if (!done.length) return;

    // Any column that parses as a date on a useful share of onboarded rows.
    var cand = [];
    raw.headers.forEach(function(h, ci) {
      if (!h) return;
      var hits = 0;
      done.forEach(function(r) { if (parseDate(r[ci])) hits++; });
      if (hits >= done.length * 0.3) cand.push({ h: h, i: ci, hits: hits });
    });
    Logger.log('\n  median span, START (row) -> END (column), in days:');
    Logger.log('    ' + '.'.repeat(28) + cand.map(function(c) {
      return (c.h.slice(0, 12) + '            ').slice(0, 13);
    }).join(''));
    cand.forEach(function(s) {
      var line = '    ' + (s.h + ' '.repeat(28)).slice(0, 28);
      cand.forEach(function(e) {
        if (s.i === e.i) { line += '           . '; return; }
        var sp = [];
        done.forEach(function(r) {
          var a = parseDate(r[s.i]), b = parseDate(r[e.i]);
          if (!a || !b) return;
          var t = dateDiffDays(a, b);
          if (t !== null && t >= 0 && t <= TAT_MAX_DAYS) sp.push(t);
        });
        sp.sort(function(x, y) { return x - y; });
        line += sp.length
          ? ((sp[Math.floor(sp.length / 2)] + 'd(' + sp.length + ')') + '            ').slice(0, 13)
          : '           - ';
      });
      Logger.log(line);
    });
    Logger.log('  (a cell is the median days from the row column to the column column,'
      + ' with the number of records in brackets)');
  });
}

// ─────────────────────────────────────────────────────────────
// Run from the Apps Script editor to confirm the buyer TAT columns. Prints the
// headers and sample values sitting at the configured positions, the span they
// produce over onboarded buyers, and a warning if the two look reversed.
// ─────────────────────────────────────────────────────────────
function debugBuyerTatCols() {
  var cfg = AUDIENCE_CFG.buyer, raw;
  try { raw = readData('buyer'); } catch (e) { Logger.log('✗ readData failed: ' + e.message); return; }
  var idx = buildIndex(raw.headers);
  var sI = BUYER_TAT_COLS.start, eI = BUYER_TAT_COLS.end;
  Logger.log('════ BUYER TAT COLUMNS ════');
  Logger.log('source: ' + raw.source + '   rows: ' + raw.rows.length + '   columns: ' + raw.headers.length);
  if (raw.headers.length <= eI) {
    Logger.log('✗ feed has only ' + raw.headers.length + ' columns — G/H are out of range.');
    return;
  }
  Logger.log('  start  col ' + colLetter_(sI) + ' = "' + (raw.headers[sI] || '(blank header)') + '"');
  Logger.log('  end    col ' + colLetter_(eI) + ' = "' + (raw.headers[eI] || '(blank header)') + '"');

  var done = raw.rows.filter(function(r) {
    return normStatus(gv(r, idx, 'onboarding_status')) === 'COMPLETED';
  });
  Logger.log('  onboarded buyers: ' + done.length);
  Logger.log('  first 5 value pairs (G → H):');
  done.slice(0, 5).forEach(function(r) {
    Logger.log('    ' + JSON.stringify(r[sI]) + '   →   ' + JSON.stringify(r[eI]));
  });

  var spans = [], neg = 0, parsedBoth = 0;
  done.forEach(function(r) {
    var a = parseDate(r[sI]), b = parseDate(r[eI]);
    if (!a || !b) return;
    parsedBoth++;
    var t = dateDiffDays(a, b);
    if (t === null) return;
    if (t < 0) neg++;
    else if (t <= TAT_MAX_DAYS) spans.push(t);
  });
  spans.sort(function(x, y) { return x - y; });
  Logger.log('  both parse as dates on ' + parsedBoth + '/' + done.length + ' onboarded buyers');
  Logger.log('  usable spans: ' + spans.length
    + (spans.length ? '  · median ' + spans[Math.floor(spans.length / 2)] + 'd'
        + ' · mean ' + Math.round(spans.reduce(function(a, b) { return a + b; }, 0) / spans.length) + 'd'
        + ' · range ' + spans[0] + '–' + spans[spans.length - 1] + 'd' : ''));
  if (neg > spans.length && neg > 0) {
    Logger.log('  ⚠ ' + neg + ' spans are NEGATIVE against ' + spans.length + ' positive —'
      + ' column G may be the END and column H the START. Swap BUYER_TAT_COLS if so.');
  }
  var rows = normalizeRows(raw, cfg);
  var withTat = rows.filter(function(r) { return r.onbTAT !== null; });
  Logger.log('  dashboard result: TAT on ' + withTat.length + ' buyers'
    + ' · avg ' + (withTat.length
        ? Math.round(withTat.reduce(function(s, r) { return s + r.onbTAT; }, 0) / withTat.length) + 'd' : '—'));

  // Attribute every onboarded buyer that did NOT get a TAT.
  var st = _lastTatStats['buyer'];
  if (st) {
    var t = st.tally;
    Logger.log('\n── why records are excluded ──');
    Logger.log('  basis chosen        : ' + (st.basis || 'none'));
    Logger.log('  candidate coverage  : G/H ' + st.coverage.fixed + ' · created ' + st.coverage.created
      + ' · review ' + st.coverage.review + ' · level1 ' + st.coverage.level1);
    Logger.log('  onboarded buyers    : ' + t.onboarded);
    Logger.log('    got a TAT         : ' + t.ok);
    Logger.log('    no start date     : ' + t.noStart   + (t.noStart ? '   <- column G empty on these rows' : ''));
    Logger.log('    no end date       : ' + t.noEnd     + (t.noEnd ? '   <- column H and every fallback empty' : ''));
    Logger.log('    start after end   : ' + t.startAfterEnd + (t.startAfterEnd ? '   <- G later than H; columns may be reversed' : ''));
    Logger.log('    span > ' + TAT_MAX_DAYS + 'd     : ' + t.tooLong);
    Logger.log('    no basis at all   : ' + t.noBasis);
    if (t.ok < t.onboarded) {
      Logger.log('  → ' + (t.onboarded - t.ok) + ' onboarded buyers have no TAT for the reasons above.');
    }
  }
}

// Map each sheet row to the common record shape the dashboard renders from.
function normalizeRows(raw, cfg) {
  var idx = buildIndex(raw.headers);

  // A dedicated TAT end column, resolved once for the feed. cfg.onbCol is the
  // onboarded date used for counts, but for buyers that is onboarding_updated_date —
  // a last-touched timestamp that moves whenever a record is edited, stretching the
  // measured span far beyond the real one. Here we look specifically for a genuine
  // completion/approval timestamp and explicitly refuse "updated"/"modified" columns.
  // The detail tab is preferred over this when it has a matching row; this covers
  // buyers, who are largely absent from the seller-centric detail card.
  var _tatEndC = tatEndColIndex_(raw);

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
    // The buyer feed names this column differently from the seller feed, so the
    // list covers both naming families; add any new header here rather than
    // letting TAT silently fall through to a different start date.
    var reviewDate = parseDate(
      gv(row, idx, 'in_review_date')              ||
      gv(row, idx, 'in_review_at')                ||
      gv(row, idx, 'under_review_date')           ||
      gv(row, idx, 'review_started_date')         ||
      gv(row, idx, 'review_started_at')           ||
      gv(row, idx, 'review_initiated_date')       ||
      gv(row, idx, 'review_date')                 ||
      gv(row, idx, 'review_at')                   ||
      gv(row, idx, 'submitted_date')              ||
      gv(row, idx, 'submitted_at')                ||
      gv(row, idx, 'submission_date')             ||
      gv(row, idx, 'application_submitted_date')  ||
      gv(row, idx, 'kyc_submitted_date')          ||
      gv(row, idx, 'onboarding_in_review_date')   ||
      gv(row, idx, 'onboarding_submitted_date')   ||
      gv(row, idx, 'level_1_date')                ||
      gv(row, idx, 'level1_date')                 ||
      gv(row, idx, 'l1_date')                     ||
      gv(row, idx, 'level_2_date')                ||
      gv(row, idx, 'level2_date')                 ||
      gv(row, idx, 'l2_date')                    || ''
    );

    // If none of the exact header names above matched, scan every column for one
    // that looks like a review/submission timestamp and actually parses as a date.
    // The exact list can never cover every feed's naming, and a missed column meant
    // TAT silently reported nothing at all.
    if (!reviewDate) {
      var _hk = Object.keys(idx);
      for (var _r = 0; _r < _hk.length; _r++) {
        var _h = _hk[_r];
        var looksReview = _h.indexOf('review') !== -1 || _h.indexOf('submit') !== -1
                       || _h.indexOf('level_1') !== -1 || _h.indexOf('level1') !== -1
                       || _h.indexOf('l1_') === 0    || _h.indexOf('_l1') !== -1;
        if (!looksReview) continue;
        // Skip columns that name a person or carry a status rather than a timestamp.
        if (_h.indexOf('_by') !== -1 || _h.indexOf('user') !== -1 || _h.indexOf('name') !== -1
            || _h.indexOf('status') !== -1 || _h.indexOf('remark') !== -1
            || _h.indexOf('comment') !== -1 || _h.indexOf('reason') !== -1
            || _h.indexOf('count') !== -1) continue;
        var _rd = parseDate(row[idx[_h]]);
        if (_rd) { reviewDate = _rd; break; }
      }
    }

    // TAT dates come from the card 5292 detail tab (_mb_detail), the dedicated
    // onboarding-stage source, matched by id then name. Its In Review date is the
    // preferred start; its Level 1 date is the fallback. The main card's own review
    // column, detected above, is used only when the detail tab has no match.
    // TAT itself is assigned after this map, in _assignTat_ — the start date is chosen
    // once for the whole feed rather than per record. Mixing starts within a feed is
    // what inflated Metal against Plastic: some records measured from review, others
    // from the earlier Level 1 date.
    var _det = lookupTatDetail_(recId, recName, cfg.audience);
    if (_det) {
      // A rejected case that came back restarts the clock. Time spent waiting for the
      // vendor to correct and resubmit is not review turnaround, so the start moves to
      // the resubmission. Where there is no explicit resubmission column, a review date
      // recorded AFTER the rejection is itself the re-review and is used; a review date
      // that predates the rejection is stale and is discarded rather than measured from.
      var _rev = _det.review, _rej = _det.rejected, _res = _det.resubmitted;
      if (_res && (!_rev || _res > _rev)) _rev = _res;
      if (_rej && _rev && _rev < _rej)    _rev = null;
      if (_rev) reviewDate = _rev;
    }
    var level1 = _det ? _det.level1 : null;
    // TAT end. The detail tab's completion date is authoritative when present and is
    // kept SEPARATE from onboardedDate, which drives onboarded counts, aging and date
    // filters and must not shift. This matters most for buyers: their card has no
    // onboarded_date, so onboardedDate falls back to onboarding_updated_date — a
    // last-touched timestamp that drifts forward every time a record is edited and
    // inflated buyer turnaround well beyond the real span.
    var _feedEnd = _tatEndC >= 0 ? parseDate(row[_tatEndC]) : null;
    var tatEnd = (_det && _det.completed) ? _det.completed
               : _feedEnd ? _feedEnd
               : onboarded;
    // Buyers: TAT comes from the buyer feed's own columns G and H, by position.
    // These take precedence over everything above for this audience.
    var _bs = null, _be = null;
    if (cfg.audience === 'buyer') {
      _bs = parseDate(row[BUYER_TAT_COLS.start]);
      _be = parseDate(row[BUYER_TAT_COLS.end]);
    }

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
      onbTAT:        null,          // assigned uniformly by _assignTat_ below
      tatBasis:      null,
      reviewDate:    reviewDate,
      _l1:           level1,        // TAT start candidate; stripped after resolution
      _tatEnd:       tatEnd,        // TAT end; separate from onboardedDate, stripped too
      _bs:           _bs,           // buyer col G start / col H end, positional
      _be:           _be,
    };
  }).filter(function(r) { return r.id || r.name; });
  _assignTat_(normalized, cfg.audience);
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
  var deduped = normalized.filter(function(r) {
    if (!r.id) return true;
    var key = r.id + '\x00' + r.vertical;
    if (seen[key]) return false;
    seen[key] = true;
    // Prefer the feed's total_orders; fall back to the pre-dedup row tally.
    r.txnCount = (orderMax[key] != null) ? orderMax[key] : (txnCounts[key] || 1);
    return true;
  });

  // Managed Marketplace = all Others rows that are NOT Transport or Support cases.
  // Others = restricted to cases whose onboarding_created_date is after 31 March 2026.
  var OTHERS_CUTOFF = new Date(2026, 3, 1); // April 1 2026
  var result = [];
  deduped.forEach(function(r) {
    if (r.vertical !== 'Others') { result.push(r); return; }
    var cat = (r.category    || '').toLowerCase();
    var bv  = (r.bizVertical || '').toLowerCase();
    var isExcluded = cat.indexOf('transport') !== -1 || bv.indexOf('transport') !== -1
                  || cat.indexOf('support')   !== -1 || bv.indexOf('support')   !== -1;
    if (!isExcluded) {
      var mmRow = {};
      for (var k in r) mmRow[k] = r[k];
      mmRow.vertical = 'Marketplace';
      result.push(mmRow);
    }
    // Keep the original Others row only if created after 31 March 2026.
    if (!r.createdDate || r.createdDate >= OTHERS_CUTOFF) {
      result.push(r);
    }
  });
  return result;
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
    if (r.onbTAT !== null) tats.push(r.onbTAT);   // onboarded-only + range enforced at assignment

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
    if (r.onbTAT !== null) cs.tats.push(r.onbTAT);
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
    var fyTats = d.filter(function(r) { return r.onbTAT !== null; })
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
    // Which start date every TAT in this vertical was measured from ('review' |
    // 'level1' | null) so the UI can label the figure honestly.
    tatBasis: (function() {
      for (var _i = 0; _i < data.length; _i++) { if (data[_i].tatBasis) return data[_i].tatBasis; }
      return null;
    }()),
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
  // v12: Vendor Rating / OSV now sourced from the "NBFC TRACKER" workbook, joined by
  // Seller_GSTIN and read by header name — bumped so no stale v10/v11 payload (old
  // source / all-unrated) is reused after this change deploys.
  var CACHE_KEY = 'quality_data_v12';
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

  var vendorRatings = [];
  var vendorOSV     = [];
  var vendorDocs    = [];

  // ══ Vendor Score workbook — Vendor Scores (col H) + OSV Status (col N) ══
  // Scores are out of 10 and bucketed into fifths for the distribution:
  // 0-2 · 2-4 · 4-6 · 6-8 · 8-10, so bucket = floor(score / 2) clamped to 0..4
  // (a score landing exactly on a boundary belongs to the higher band, e.g. 4 → 4-6,
  // and a perfect 10 folds back into the top band).
  // Rows join to onboarded vendors by GSTIN first, then by id — same keys the rest
  // of the quality build uses. Vendors absent from this workbook stay unrated.
  try {
    var vsd = qualityReadExternalSheet_(CONFIG.VENDOR_SCORE_SHEET_ID, CONFIG.VENDOR_SCORE_TAB);
    if (vsd.error) {
      Logger.log('Vendor Score workbook unreadable (' + vsd.error + ') — rating/OSV will be empty.');
    } else if (vsd.rows.length) {
      var vsh   = vsd.headers;
      var vsGst = qualityFindCol_(vsh, ['seller_gstin','buyer_gstin','gstin','gst_number','gst_no','gstin_number','gst']);
      var vsId  = qualityFindCol_(vsh, ['seller_id','buyer_id','id','vendor_id','vendorid','entity_id']);
      var vsNm  = qualityFindCol_(vsh, ['seller_name','name','vendor_name','business_name','company_name']);
      // Optional enrichment for the OSV records table. Absent columns simply leave
      // the corresponding cell blank rather than blocking the row.
      var vsCat = qualityFindCol_(vsh, ['business_category','category','vertical_category','cat','material','material_type']);
      var vsBv  = qualityFindCol_(vsh, ['business_vertical','vertical','biz_vertical']);
      var vsAsg = qualityFindCol_(vsh, ['assigned_user','assigned_to','assignee','owner','sales_poc','poc',
                                        'account_manager','relationship_manager','rm','kam','executive','agent']);
      var vsUpd = qualityFindCol_(vsh, ['last_updated','last_updated_date','last_updated_at','updated_at',
                                        'updated_date','modified_at','last_modified','last_modified_date',
                                        'osv_date','osv_updated_date','consent_date','status_date']);

      // The join keys (gstin / id) are resolved by header NAME, but the score and OSV
      // were read by fixed POSITION (col H / col N). If the Vendor Score sheet's columns
      // get reordered, that positional read silently lands on the wrong cell — the join
      // still matches the vendor, but the "score" parses as NaN, so EVERY vendor counts
      // as unrated and the module shows 0 rated / N unrated (exactly this failure).
      // Resolve both by header name first and fall back to the configured position only
      // when no recognisable header is present, so a column reorder can't zero it out.
      var vsScore = qualityFindCol_(vsh, ['finoscale_rating','finoscale_score','finoscale','vendor_score','vendor_scores','vendorscore','overall_score',
        'average_score','avg_score','final_score','score_out_of_10','vendor_rating_score','quality_score','seller_score','score']);
      if (vsScore < 0) vsScore = VS_COLS.denominator;
      var vsOsv = qualityFindCol_(vsh, ['osv_status','osv','on_site_verification','onsite_verification',
        'osv_state','osv_consent_status','consent_status','verification_status','osv_verification_status']);
      if (vsOsv < 0) vsOsv = VS_COLS.osvStatus;

      // GSTIN is a case-insensitive identifier; normalise (upper-case, strip spaces &
      // punctuation) on both sides so a formatting drift in the sheet can't break the join.
      var _nrmG = function(g) { return String(g || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); };
      var ompGstinNorm = {};
      Object.keys(ompGstinMap).forEach(function(g) { ompGstinNorm[_nrmG(g)] = ompGstinMap[g]; });

      vsd.rows.forEach(function(row) {
        var gst = vsGst >= 0 ? String(row[vsGst] || '').trim() : '';
        var vid = vsId  >= 0 ? String(row[vsId]  || '').trim() : '';
        var omp = (gst && (ompGstinMap[gst] || ompGstinNorm[_nrmG(gst)])) || (vid && ompMap[vid]) || null;
        if (!omp) return;   // not an OMP-onboarded vendor — out of scope
        var aud = omp.aud === 'buyer' ? 'buyer' : 'seller';

        // — Vendor Score (out of 10; column resolved above) —
        // A blank cell is unrated (NaN) and simply doesn't count. A literal 0 is a
        // real score and is kept, so it lands in the 0-2 band rather than vanishing.
        var sv = parseFloat(row[vsScore]);
        if (!isNaN(sv) && sv >= 0 && sv <= 10) {
          var band = Math.min(4, Math.max(0, Math.floor(sv / 2)));
          [aud, 'combined'].forEach(function(t) {
            acc[t].r.ws    += sv;
            acc[t].r.rated += 1;
            acc[t].r.dist[band] += 1;
          });
          vendorRatings.push({
            id:               omp.id   || vid.slice(0, 30),
            name:             omp.name || (vsNm >= 0 ? String(row[vsNm] || '').trim().slice(0, 60) : ''),
            gstin:            (omp.gstin || gst).slice(0, 20),
            rating:           Math.round(sv * 10) / 10,
            onboardingStatus: omp.onboardingStatus,
            category:         omp.category,
            aud:              aud
          });
        }

        // — OSV Status (column N; three-way: verified / in-progress / not-initiated) —
        // Sellers only, matching how OSV is scoped everywhere else in the dashboard.
        if (aud === 'seller') {
          var osRaw = String(row[vsOsv] || '').trim();
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
            id:               (omp.id || vid).slice(0, 30),
            name:             omp.name || (vsNm >= 0 ? String(row[vsNm] || '').trim().slice(0, 60) : ''),
            gstin:            omp.gstin || gst.slice(0, 20),
            consentRaw:       osRaw.slice(0, 30),
            status:           osvStatus,
            onboardingStatus: omp.onboardingStatus,
            // Prefer the onboarding record's category/vertical (authoritative); fall
            // back to the score sheet's own columns when the join carries neither.
            category:         omp.category || (vsCat >= 0 ? String(row[vsCat] || '').trim().slice(0, 60) : ''),
            vertical:         omp.bizVertical || (vsBv >= 0 ? String(row[vsBv] || '').trim().slice(0, 40) : ''),
            assignedTo:       vsAsg >= 0 ? String(row[vsAsg] || '').trim().slice(0, 60) : '',
            updatedDate:      (function() {
                                if (vsUpd < 0) return '';
                                var _d = parseDate(row[vsUpd]);
                                return _d ? fmtDate(_d) : '';   // '' not '—', so exports stay clean
                              }()),
            aud:              'seller'
          });
        }
      });
    }
  } catch (e) { Logger.log('getQualityData vendor-score sheet: ' + e.message); }

  // ── "Vendor Rating" sheet — document completeness only. Rating and OSV now
  //    come from the Vendor Score workbook above.
  // Cols 32-42 = per-document 0/1 flags (11 documents)
  try {
    var vrd = qualityReadSheet_('Vendor Rating');
    if (vrd.rows.length) {
      var vrh = vrd.headers;

      var vidC  = qualityFindCol_(vrh, ['seller_id','id','vendor_id','vendorid']);
      var vauC  = qualityFindCol_(vrh, ['audience','type','aud','seller_buyer']);
      var nameC = qualityFindCol_(vrh, ['seller_name','name','vendor_name','business_name','company_name']);
      if (nameC < 0) nameC = 1;

      // GSTIN column, used to enrich document rows with OMP entity info
      var vrGstC = qualityFindCol_(vrh, ['gstin','gst_number','gst_no','gstin_number','gst']);

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

      vendorDocs = [];

      vrd.rows.forEach(function(row) {
        var rowAud = resolveAud_(row, vrh, vidC, vauC);
        var ts = tgts_(rowAud);

        // Rating and OSV are no longer read here — both now come from the Vendor
        // Score workbook (see the dedicated pass above). This sheet supplies
        // document completeness only.

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
      // Raw business vertical, kept for the OSV records table. Always an Open
      // Marketplace value here — the filter above admits nothing else.
      bizVertical:      bvC  >= 0 ? String(row[bvC]  || '').trim().slice(0, 40) : '',
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

// Read a tab from an EXTERNAL workbook by file id. Same shape as
// qualityReadSheet_ (normalized headers + non-empty rows) but tolerant: any
// failure (bad id, no access, missing tab) returns an empty result with the
// reason attached, so a broken external source degrades to the in-workbook
// fallback instead of throwing the whole quality build.
function qualityReadExternalSheet_(sheetId, tabName) {
  if (!sheetId) return { headers: [], rows: [], error: 'no sheet id configured' };
  try {
    var ss    = SpreadsheetApp.openById(sheetId);
    var sheet = tabName ? ss.getSheetByName(tabName) : ss.getSheets()[0];
    if (!sheet) return { headers: [], rows: [], error: 'tab "' + tabName + '" not found' };
    if (sheet.getLastRow() < 2) return { headers: [], rows: [], error: 'sheet has no data rows' };
    var vals    = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    var headers = vals[0].map(function(h) {
      return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
    });
    var rows = vals.slice(1).filter(function(row) {
      return row.some(function(c) { return c !== '' && c !== null && c !== undefined; });
    });
    return { headers: headers, rows: rows, tab: sheet.getName() };
  } catch (e) {
    return { headers: [], rows: [], error: e.message };
  }
}

// 0-based column index → spreadsheet column letter (7 → "H", 13 → "N").
function colLetter_(i) {
  var s = '', n = i + 1;
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ─────────────────────────────────────────────────────────────
// Run manually from the Apps Script editor to inspect the Vendor Score
// workbook: dumps every header with its column letter, flags what sits at the
// configured positions (H / N), and prints sample values so the mapping can be
// confirmed or corrected before trusting the numbers.
// ─────────────────────────────────────────────────────────────
function debugVendorScoreSheet() {
  var d = qualityReadExternalSheet_(CONFIG.VENDOR_SCORE_SHEET_ID, CONFIG.VENDOR_SCORE_TAB);
  Logger.log('Vendor Score workbook: ' + CONFIG.VENDOR_SCORE_SHEET_ID);
  if (d.error) { Logger.log('✗ could not read: ' + d.error); return; }
  Logger.log('tab: "' + d.tab + '"   rows: ' + d.rows.length + '   columns: ' + d.headers.length);
  Logger.log('\n── all columns ──');
  d.headers.forEach(function(h, i) {
    Logger.log('  ' + colLetter_(i) + ' (idx ' + i + '): ' + (h || '(blank header)'));
  });

  [['denominator', VS_COLS.denominator], ['osvStatus', VS_COLS.osvStatus]].forEach(function(p) {
    var key = p[0], ci = p[1];
    Logger.log('\n── configured ' + key + ' → column ' + colLetter_(ci) + ' ──');
    if (ci >= d.headers.length) { Logger.log('  ✗ column out of range — sheet has only ' + d.headers.length + ' columns'); return; }
    Logger.log('  header: ' + (d.headers[ci] || '(blank)'));
    var vals = d.rows.slice(0, 8).map(function(r) { return JSON.stringify(r[ci]); });
    Logger.log('  first 8 values: ' + vals.join(', '));
    var nums = d.rows.map(function(r) { return parseFloat(r[ci]); }).filter(function(n) { return !isNaN(n); });
    if (nums.length) {
      Logger.log('  numeric in ' + nums.length + '/' + d.rows.length + ' rows · min ' + Math.min.apply(null, nums)
        + ' · max ' + Math.max.apply(null, nums));
    } else {
      var distinct = {};
      d.rows.forEach(function(r) { var v = String(r[ci] || '').trim(); if (v) distinct[v] = (distinct[v] || 0) + 1; });
      Logger.log('  non-numeric · distinct values: ' + JSON.stringify(distinct).slice(0, 500));
    }
  });

  // Which columns could join these rows back to the dashboard's vendors.
  var idC  = qualityFindCol_(d.headers, ['seller_id','buyer_id','id','vendor_id','vendorid','entity_id']);
  var gstC = qualityFindCol_(d.headers, ['seller_gstin','buyer_gstin','gstin','gst_number','gst_no','gstin_number','gst']);
  Logger.log('\n── join keys ──');
  Logger.log('  id column:    ' + (idC  >= 0 ? colLetter_(idC)  + ' (' + d.headers[idC]  + ')' : '✗ NOT FOUND'));
  Logger.log('  gstin column: ' + (gstC >= 0 ? colLetter_(gstC) + ' (' + d.headers[gstC] + ')' : '✗ NOT FOUND'));
  if (idC < 0 && gstC < 0) Logger.log('  ⚠ no join key detected — rows cannot be matched to onboarded vendors.');
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
