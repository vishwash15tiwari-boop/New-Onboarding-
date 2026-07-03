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
  SELLER_SHEET_ID: '1qaG_GMvUrC7LJbKBma8-S3x2N6Ua4vDkUC5ZRx3znho',
  BUYER_SHEET_ID:  '1G9Ocq8PXovCx5eBE3dOfE8CUcmfgwcl6Te37p79u0wI',
  CACHE_TTL: 55,   // seconds — < the 60 s frontend poll so each refresh is fresh
};

// Per-audience column mapping (the two sheets name a few fields differently).
var AUDIENCE_CFG = {
  seller: { sheetId: CONFIG.SELLER_SHEET_ID, label: 'Sellers',
            idCol: 'seller_id', nameCol: 'seller_name', vendorCol: 'vendor_type', onbCol: 'onboarded_date' },
  // Buyers sheet has no onboarded_date, so TAT uses onboarding_updated_date
  // (the last status change) as the onboarding-complete proxy.
  buyer:  { sheetId: CONFIG.BUYER_SHEET_ID,  label: 'Buyers',
            idCol: 'buyer_id',  nameCol: 'buyer_name',  vendorCol: 'customer_type', onbCol: 'onboarding_updated_date' },
};

// TAT beyond this many days is treated as data noise (bulk re-import timestamps)
// and excluded from the Avg TAT so a few outliers don't skew it.
var TAT_MAX_DAYS = 365;

// The seven business verticals, in display order. Rows are mapped into these by
// mapToVertical(); every audience always shows all seven.
var VERTICALS = [
  { key: 'OMP',           name: 'Open Marketplace', code: 'OMP', sub: 'Open Marketplace onboarding' },
  { key: 'EPR',           name: 'EPR',              code: 'EPR', sub: 'Extended Producer Responsibility' },
  { key: 'InfraBusiness', name: 'Infra Business',   code: 'INF', sub: 'Metal · Plastic · Institutional · Reverse' },
  { key: 'AFR',           name: 'AFR',              code: 'AFR', sub: 'Alternative Fuels & Resources' },
  { key: 'Recommerce',    name: 'Re-Commerce',      code: 'REC', sub: 'Re-Commerce · E-Waste' },
  { key: 'DRS',           name: 'DRS',              code: 'DRS', sub: 'Deposit Refund System' },
  { key: 'Others',        name: 'Others',           code: 'OTH', sub: 'Other verticals & categories' },
];

// Categories (from the Marketplace business_vertical) that roll into Infra Business.
var INFRA_CATS = { 'metal': 1, 'plastic': 1, 'institutional business': 1, 'reverse': 1, 'rewerse': 1 };
function isEwaste(cat) { return cat === 'e-waste' || cat === 'ewaste' || cat === 'e waste'; }

// Map a sheet row to one of the seven verticals. AFR / GOA DRS / Re-Commerce are
// business_CATEGORY values under the Marketplace vertical; EPR and Open Marketplace
// are business_VERTICAL values.
//  · EPR              ← business_vertical 'EPR' (all of it)
//  · Open Marketplace ← business_vertical 'Open Marketplace', categories Metal & Plastic
//  · Marketplace vertical, by business_category:
//        AFR                                          → AFR
//        GOA DRS                                      → DRS
//        Re-Commerce, E-Waste                         → Re-Commerce
//        Metal, Plastic, Institutional Business, Reverse → Infra Business
//        anything else                                → Others
//  · everything else (Support, Sustainability Services, blank, …) → Others
function mapToVertical(businessVertical, category) {
  var bv  = String(businessVertical || '').trim().toLowerCase();
  var cat = String(category || '').trim().toLowerCase();
  if (bv === 'epr') return 'EPR';
  if (bv === 'open marketplace' || bv === 'openmarketplace') {
    return (cat === 'metal' || cat === 'plastic') ? 'OMP' : 'Others';
  }
  if (bv === 'marketplace') {
    if (cat === 'afr') return 'AFR';
    if (cat === 'goa drs' || cat === 'drs') return 'DRS';
    if (cat === 're-commerce' || cat === 'recommerce' || cat === 're commerce' || isEwaste(cat)) return 'Recommerce';
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

    var cacheKey = 'dash_v6_' + audience + '_' + JSON.stringify([f.period || 'All', f.startDate || '', f.endDate || '']);
    var cache = CacheService.getScriptCache();
    var hit = cache.get(cacheKey);
    if (hit) return hit;

    var raw  = readSheet(cfg.sheetId);
    var rows = normalizeRows(raw, cfg);
    var out  = JSON.stringify(buildDashboard(audience, cfg, rows, f));
    try { cache.put(cacheKey, out, CONFIG.CACHE_TTL); } catch (e) {}
    return out;
  } catch (err) {
    return JSON.stringify({ success: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// ─────────────────────────────────────────────────────────────
// SHEET READ + NORMALIZATION
// ─────────────────────────────────────────────────────────────
function readSheet(sheetId) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  var vals  = sheet.getDataRange().getValues();
  if (!vals.length) return { headers: [], rows: [] };
  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });
  var rows = vals.slice(1).filter(function(row) {
    return row.some(function(c) { return c !== '' && c !== null && c !== undefined; });
  });
  return { headers: headers, rows: rows };
}

// Map each sheet row to the common record shape the dashboard renders from.
function normalizeRows(raw, cfg) {
  var idx = buildIndex(raw.headers);
  return raw.rows.map(function(row) {
    var status   = normStatus(gv(row, idx, 'onboarding_status'));
    var category = normCategory(gv(row, idx, 'business_category'));
    var bizVert  = gv(row, idx, 'business_vertical');
    var created  = parseDate(gv(row, idx, 'onboarding_created_date'));
    var onboarded = cfg.onbCol ? parseDate(gv(row, idx, cfg.onbCol)) : null;
    // TAT = whole days from created to onboarded, for completed rows that have an
    // onboarded_date. dateDiffDays() compares calendar dates only, so same-day
    // onboarding (created/onboarded differ only by timezone) reads 0, not negative.
    var tat = (status === 'COMPLETED' && created && onboarded) ? dateDiffDays(created, onboarded) : null;

    var gstin     = String(gv(row, idx, 'gstin') || '').trim();
    var gstStatus = String(gv(row, idx, 'gstin_status') || '').toUpperCase();
    var txn       = String(gv(row, idx, 'transaction_activation_status') || '').toUpperCase().replace(/\s+/g, ' ').trim();
    return {
      id:            String(gv(row, idx, cfg.idCol) || '').replace(/,/g, '').trim(),
      name:          String(gv(row, idx, cfg.nameCol) || '').trim(),
      state:         String(gv(row, idx, 'state') || '').trim(),
      vendorType:    String(gv(row, idx, cfg.vendorCol) || '').trim(),
      category:      category,
      vertical:      mapToVertical(bizVert, category),
      gstin:         gstin,
      status:        status,
      createdDate:   created,
      onboardedDate: onboarded,
      // gstin_status is authoritative — real values are "GSTIN AVAILABLE" vs
      // "MISSING GSTIN" (older sheets used "GSTIN NOT AVAILABLE"). GST is active
      // only when AVAILABLE and not NOT/MISSING.
      hasGST:        gstStatus ? (gstStatus.indexOf('AVAILABLE') !== -1 && gstStatus.indexOf('NOT') === -1 && gstStatus.indexOf('MISSING') === -1)
                               : isValidGSTIN(gstin),
      hasTxn:        txn !== '',            // does this row carry a transaction signal at all?
      hasTransacted: txn === 'TRANSACTED',
      onbTAT:        tat,
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD ASSEMBLY
// ─────────────────────────────────────────────────────────────
function buildDashboard(audience, cfg, allRows, f) {
  var rows = allRows.filter(function(r) { return applyDateFilter(r, f); });

  var byVert = {};
  rows.forEach(function(r) { (byVert[r.vertical] = byVert[r.vertical] || []).push(r); });

  // Always emit all six verticals, in fixed order (empty ones render "No records").
  var verticalConfig = VERTICALS.map(function(vc) {
    return { key: vc.key, name: vc.name, sub: vc.sub, code: vc.code };
  });
  var verticals = {};
  VERTICALS.forEach(function(vc) {
    var s = vStats(byVert[vc.key] || []);
    // Transacted is only meaningful for Open Marketplace; hide it elsewhere.
    if (vc.key !== 'OMP') { s.transacted = null; s.pctTransacted = null; }
    verticals[vc.key] = s;
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

// Per-vertical KPIs + onboarded-by-category breakdown + detail rows.
function vStats(data) {
  var now = new Date(), weekAgo = new Date(now.getTime() - 7 * 86400000);
  var total     = data.length;
  var completed = count(data, 'status', 'COMPLETED');
  // GST Active/Inactive are counted over onboarded (completed) cases only.
  var withGST   = countFn(data, function(r) { return r.status === 'COMPLETED' && r.hasGST; });
  var tats      = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0 && r.onbTAT <= TAT_MAX_DAYS; }).map(function(r) { return r.onbTAT; });
  // Transacted is null (shown as "—") for a vertical whose rows carry no
  // transaction signal at all (e.g. EPR); otherwise it's the TRANSACTED count.
  var transacted = countFn(data, function(r) { return r.hasTxn; }) ? countFn(data, function(r) { return r.hasTransacted; }) : null;

  var catMap = {};
  data.forEach(function(r) {
    var c = r.category || 'Others';
    if (!catMap[c]) catMap[c] = { category: c, onboarded: 0, total: 0 };
    catMap[c].total++;
    if (r.status === 'COMPLETED') catMap[c].onboarded++;
  });
  var categories = Object.keys(catMap).map(function(k) { return catMap[k]; })
    .sort(function(a, b) { return b.onboarded - a.onboarded || b.total - a.total; });

  var latestFirst = data.slice().sort(function(a, b) {
    return (b.createdDate ? b.createdDate.getTime() : 0) - (a.createdDate ? a.createdDate.getTime() : 0);
  });

  return {
    total: total,
    completed: completed,
    draft: count(data, 'status', 'DRAFT'),
    inReview: count(data, 'status', 'IN_REVIEW'),
    rejected: count(data, 'status', 'REJECTED'),
    withGST: withGST,
    missingGST: completed - withGST,
    completionPct: pct(completed, total),
    completedThisWeek: countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= weekAgo; }),
    avgTAT: tats.length ? Math.round(avg(tats)) : null,
    transacted: transacted,
    pctTransacted: transacted === null ? null : pct(transacted, completed),
    categories: categories,
    rows: latestFirst.map(vertRow),
    rowsTotal: total,
  };
}

function vertRow(r) {
  return {
    id: r.id, name: r.name, category: r.category, vendorType: r.vendorType,
    status: r.status, gstin: r.gstin, hasGST: r.hasGST, state: r.state,
    createdDate: fmtDate(r.createdDate), onbDate: fmtDate(r.onboardedDate),
    tat: (r.onbTAT === null ? '—' : r.onbTAT),
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
    if (d) { var y = fyStartYear(d); if (y >= 2015 && y < minFY) minFY = y; }
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
