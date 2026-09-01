// ════════════════════════════════════════════════════════════════
// METABASE → ONBOARDING DASHBOARD SYNC          Metabase.gs
//
// Pulls live Seller + Buyer onboarding data from meta.recykal.com
// into hidden local sheets every 5 minutes.  Code.gs reads those
// local sheets first; if they're empty it falls back to the original
// external Google Sheets — so the dashboard degrades gracefully if
// the sync hasn't run yet.
//
// Cards:
//   5712  Seller Info              → _mb_sellers sheet
//   5711  Buyer from Inspection    → _mb_buyers  sheet
//   5292  Onboarding Detail        → _mb_detail  sheet
//
// SETUP (one time, in this order):
//   1. Paste this file into your Apps Script project (beside Code.gs)
//   2. Run  setupMetabaseCredentials()  — enter password, run, blank it out
//   3. Run  syncAllOnboarding()         — confirm sheets appear + are populated
//   4. Run  setupMetabaseTrigger()      — installs 5-min auto-sync
// ════════════════════════════════════════════════════════════════

// ── Card IDs ─────────────────────────────────────────────────────
var MB_CARDS = {
  seller: 5712,   // Seller Info
  buyer:  5711,   // Buyer from Inspection
  detail: 5292,   // Onboarding Detail (enriched)
};

// ── Quality / compliance card IDs (OMP-specific metrics) ─────────
var MB_CARDS_QUALITY = {
  rating: 5662,   // Vendor Rating
  osv:    5670,   // OSV (On-Site Verification) Status
  docs:   5674,   // Document Completeness
};

// ── Local sheet names (created hidden in this spreadsheet) ────────
var MB_SHEETS = {
  seller: '_mb_sellers',
  buyer:  '_mb_buyers',
  detail: '_mb_detail',
};

// ── Visible sheet names for quality data (shown as spreadsheet tabs)
var MB_SHEETS_QUALITY = {
  rating: 'Vendor Rating',
  osv:    'OSV Status',
  docs:   'Doc Completeness',
};

var MB_HOST = 'https://meta.recykal.com';

// Session token is cached in Script Properties for up to 12 h so
// every sync call does not re-authenticate.
var MB_TOKEN_TTL_MS = 12 * 60 * 60 * 1000;


// ════════════════════════════════════════════════════════════════
// ONE-TIME CREDENTIAL SETUP
// Run ONCE → credentials saved in Script Properties → blank the
// MB_PASSWORD line → save.  Credentials are never hard-coded.
// ════════════════════════════════════════════════════════════════
function setupMetabaseCredentials() {
  // Metabase: https://meta.recykal.com  (host is set in MB_HOST above)
  // User:     vishwash.tiwari@recykal.com
  // Paste the password on the line below, run this ONCE, then blank it again. It is
  // stored in Script Properties (never hard-coded). A BLANK value is ignored, so
  // re-running this later — e.g. to reconfirm the email — never wipes a saved password.
  var MB_PASSWORD = '';   // ← paste password here, run once, then blank this
  var props = { 'MB_EMAIL': 'vishwash.tiwari@recykal.com' };
  if (MB_PASSWORD) props['MB_PASSWORD'] = MB_PASSWORD;
  PropertiesService.getScriptProperties().setProperties(props);
  Logger.log(MB_PASSWORD
    ? '✅ Email + password saved to Script Properties.'
    : '✅ Email saved. Now paste the password on the MB_PASSWORD line, run once, then blank it.');
}


// ════════════════════════════════════════════════════════════════
// AUTHENTICATION — cached session token
// ════════════════════════════════════════════════════════════════
function getMBToken_() {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('MB_TOKEN');
  var ts     = parseInt(props.getProperty('MB_TOKEN_TS') || '0', 10);

  if (token && (Date.now() - ts) < MB_TOKEN_TTL_MS) return token;

  var email = props.getProperty('MB_EMAIL');
  var pass  = props.getProperty('MB_PASSWORD');
  if (!email || !pass) {
    throw new Error('Metabase credentials missing — run setupMetabaseCredentials() first.');
  }

  var res = UrlFetchApp.fetch(MB_HOST + '/api/session', {
    method:             'POST',
    contentType:        'application/json',
    payload:            JSON.stringify({ username: email, password: pass }),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Metabase auth failed (' + res.getResponseCode() + '): ' +
      res.getContentText().slice(0, 300));
  }

  token = JSON.parse(res.getContentText()).id;
  props.setProperties({ 'MB_TOKEN': token, 'MB_TOKEN_TS': String(Date.now()) });
  Logger.log('  Auth: new session token obtained.');
  return token;
}

function clearMBToken_() {
  PropertiesService.getScriptProperties().deleteProperty('MB_TOKEN');
}


// ════════════════════════════════════════════════════════════════
// CORE FETCH — Metabase card → { headers, rows }
// Returns the same shape as Code.gs readSheet() so normalizeRows()
// works on it directly.  Headers are lowercased + underscored to
// match the existing column-name convention.
// Retries: 401 → refresh the session token and retry; a transient failure (network
// error / timeout / 429 / 5xx) → short backoff and retry, up to 3 attempts. A large
// card (e.g. Sellers) occasionally times out or 5xx's; without a retry its sheet is
// left EMPTY until the next 5-min sync, which is exactly the "seller data disappears"
// symptom. No artificial short deadline — let a big query finish.
// ════════════════════════════════════════════════════════════════
function fetchMBCard_(cardId) {
  var token = getMBToken_();

  function doFetch_(t) {
    return UrlFetchApp.fetch(MB_HOST + '/api/card/' + cardId + '/query/csv', {
      method:             'POST',
      headers:            { 'X-Metabase-Session': t },
      contentType:        'application/json',
      payload:            JSON.stringify({}),
      muteHttpExceptions: true,
    });
  }

  var res = null, code = 0, lastErr = '';
  for (var attempt = 1; attempt <= 3; attempt++) {
    try { res = doFetch_(token); code = res.getResponseCode(); }
    catch (e) { lastErr = e.message; code = 0; res = null; }   // network/timeout throws

    if (code === 200) break;
    if (code === 401) { clearMBToken_(); token = getMBToken_(); continue; }  // refresh + retry now
    if (attempt < 3 && (code === 0 || code === 429 || code >= 500)) {
      Utilities.sleep(1500 * attempt);   // transient — back off, then retry
      continue;
    }
    break;   // non-retryable (other 4xx) or attempts exhausted
  }

  if (code !== 200) {
    throw new Error('Card ' + cardId + ' fetch failed (' + (code || 'network: ' + lastErr) + ')' +
      (res ? ': ' + res.getContentText().slice(0, 300) : ''));
  }

  var csv = Utilities.parseCsv(res.getContentText());
  if (!csv || csv.length < 2) {
    Logger.log('  ⚠ Card ' + cardId + ' returned no data rows.');
    return { headers: [], rows: [] };
  }

  // Normalise headers identically to Code.gs readSheet()
  var headers = csv[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
  });

  return { headers: headers, rows: csv.slice(1) };
}


// ════════════════════════════════════════════════════════════════
// WRITE — flush fetched data into a local hidden sheet
// Only data rows are erased (row 2+) so the header is always
// visible during a write.  Sync metadata goes to Script Properties,
// not to the sheet, so it never leaks into the dashboard data.
// ════════════════════════════════════════════════════════════════
function writeMBSheetVisible_(data, sheetName) {
  if (!data.headers.length) return;
  var payload = [data.headers].concat(data.rows);
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);   // NOT hidden — visible as a spreadsheet tab
    Logger.log('  Created visible sheet "' + sheetName + '".');
  }
  var usedRows = sheet.getLastRow();
  var usedCols = Math.max(sheet.getLastColumn(), payload[0].length);
  // Write the new data FIRST (overwrites in place), THEN clear only the leftover
  // trailing rows. The old order (clearContent → setValues) left the sheet EMPTY if
  // setValues failed/timed out mid-write — which is the "_mb_* is empty, sellers
  // vanish" failure. Write-then-trim keeps the previous good data on any write failure.
  sheet.getRange(1, 1, payload.length, payload[0].length).setValues(payload);
  if (usedRows > payload.length) {
    sheet.getRange(payload.length + 1, 1, usedRows - payload.length, usedCols).clearContent();
  }
  PropertiesService.getScriptProperties()
    .setProperty('MB_SYNC_' + sheetName, new Date().toISOString());
}

function writeMBSheet_(data, sheetName) {
  if (!data.headers.length) return;

  // Reassemble as 2-D array: header row + data rows
  var payload = [data.headers].concat(data.rows);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();   // helper sheet — not for end users
    Logger.log('  Created hidden sheet "' + sheetName + '".');
  }

  // Write header + data FIRST (overwrites in place), THEN clear only the leftover
  // trailing rows. The old order (clearContent → setValues) left the sheet EMPTY when
  // setValues failed/timed out mid-write — the "_mb_sellers is empty, sellers vanish
  // from the dashboard" failure. Write-then-trim keeps the last good data on failure.
  var usedRows = sheet.getLastRow();
  var usedCols = Math.max(sheet.getLastColumn(), payload[0].length);
  sheet.getRange(1, 1, payload.length, payload[0].length).setValues(payload);
  if (usedRows > payload.length) {
    sheet.getRange(payload.length + 1, 1, usedRows - payload.length, usedCols).clearContent();
  }

  // Record sync timestamp in Script Properties (not in the sheet — avoids
  // polluting the data range that Code.gs reads).
  PropertiesService.getScriptProperties()
    .setProperty('MB_SYNC_' + sheetName, new Date().toISOString());
}


// ════════════════════════════════════════════════════════════════
// CACHE BUST — invalidate Code.gs dashboard cache so the next
// frontend load picks up the freshly synced data immediately.
// ════════════════════════════════════════════════════════════════
function bustDashboardCache_() {
  var cache = CacheService.getScriptCache();
  var keys  = [];
  var periods = ['All', 'Today', 'ThisWeek', 'ThisMonth', 'MTD', 'YTD'];

  // Add FY keys for the last 5 years
  var nowYear = new Date().getFullYear();
  for (var y = nowYear - 4; y <= nowYear; y++) {
    periods.push('FY' + String(y).slice(2) + '-' + String(y + 1).slice(2));
  }

  var vertKeys = ['OMP', 'EPR', 'Marketplace', 'InfraBusiness', 'AFR', 'Recommerce', 'DRS', 'Others'];

  // Version ranges must COVER the current keys used in Code.gs (dash_v36 / vrows_v17)
  // plus a forward margin — the previous list only cleared v30-v32 / v14-v16, so the
  // live keys were never invalidated and every sync left up-to-5-min-old data in cache
  // (the root of "data not updating / not syncing"). Ranges auto-cover a version bump.
  var dashPfx = [], vrowsPfx = [], cmbPfx = [];
  for (var dv = 30; dv <= 42; dv++) { dashPfx.push('dash_v' + dv + '_'); cmbPfx.push('dash_v' + dv + '_cmb_'); }
  for (var rv = 14; rv <= 22; rv++) { vrowsPfx.push('vrows_v' + rv + '_'); }

  ['seller', 'buyer'].forEach(function(aud) {
    periods.forEach(function(p) {
      var pk = JSON.stringify([p, '', '']);
      dashPfx.forEach(function(pfx) { keys.push(pfx + aud + '_' + pk); });
      vrowsPfx.forEach(function(pfx) {
        vertKeys.forEach(function(vk) { keys.push(pfx + aud + '_' + vk + '_' + pk); });
      });
    });
  });

  // Combined dashboard cache (getCombinedDashboard) — same version range.
  periods.forEach(function(p) {
    var pk = JSON.stringify([p, '', '']);
    cmbPfx.forEach(function(pfx) { keys.push(pfx + pk); });
  });

  // Quality metrics cache — range covering the current key (quality_data_v13) + margin.
  for (var qv = 2; qv <= 20; qv++) keys.push('quality_data_v' + qv);

  // Batch the removals (chunked removeAll) — this list is ~2k keys and per-key remove()
  // would be one RPC each, slow enough to matter inside the 5-min sync run.
  for (var ci = 0; ci < keys.length; ci += 200) {
    try { cache.removeAll(keys.slice(ci, ci + 200)); } catch (e) {}
  }
  Logger.log('  Cache busted (' + keys.length + ' keys).');
}


// ════════════════════════════════════════════════════════════════
// SYNC FUNCTIONS — one per card + master entry point
// ════════════════════════════════════════════════════════════════
function syncSellers() {
  Logger.log('▶ Syncing Sellers (card ' + MB_CARDS.seller + ')…');
  var data = fetchMBCard_(MB_CARDS.seller);
  writeMBSheet_(data, MB_SHEETS.seller);
  Logger.log('✅ Sellers: ' + data.rows.length + ' rows → "' + MB_SHEETS.seller + '"');
}

function syncBuyers() {
  Logger.log('▶ Syncing Buyers (card ' + MB_CARDS.buyer + ')…');
  var data = fetchMBCard_(MB_CARDS.buyer);
  writeMBSheet_(data, MB_SHEETS.buyer);
  Logger.log('✅ Buyers: ' + data.rows.length + ' rows → "' + MB_SHEETS.buyer + '"');
}

// Card 5292 is the dedicated TAT source (Code.gs getTATLookup_ reads its
// "In Review" column). Kept as a VISIBLE tab so the imported query results can
// be inspected and audited, and refreshed on every 5-min sync.
function syncDetail() {
  Logger.log('▶ Syncing Onboarding Detail (card ' + MB_CARDS.detail + ')…');
  var data = fetchMBCard_(MB_CARDS.detail);
  writeMBSheetVisible_(data, MB_SHEETS.detail);
  // Un-hide if an earlier build created it as a hidden helper sheet.
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MB_SHEETS.detail);
    if (sh && sh.isSheetHidden()) sh.showSheet();
  } catch (e) { /* non-fatal */ }
  Logger.log('✅ Detail: ' + data.rows.length + ' rows → "' + MB_SHEETS.detail + '" (TAT source)');
}

function syncRating() {
  Logger.log('▶ Syncing Vendor Rating (card ' + MB_CARDS_QUALITY.rating + ')…');
  var data = fetchMBCard_(MB_CARDS_QUALITY.rating);
  writeMBSheetVisible_(data, MB_SHEETS_QUALITY.rating);
  Logger.log('✅ Rating: ' + data.rows.length + ' rows → "' + MB_SHEETS_QUALITY.rating + '"');
}

function syncOSV() {
  Logger.log('▶ Syncing OSV Status (card ' + MB_CARDS_QUALITY.osv + ')…');
  var data = fetchMBCard_(MB_CARDS_QUALITY.osv);
  writeMBSheetVisible_(data, MB_SHEETS_QUALITY.osv);
  Logger.log('✅ OSV: ' + data.rows.length + ' rows → "' + MB_SHEETS_QUALITY.osv + '"');
}

function syncDocs() {
  Logger.log('▶ Syncing Doc Completeness (card ' + MB_CARDS_QUALITY.docs + ')…');
  var data = fetchMBCard_(MB_CARDS_QUALITY.docs);
  writeMBSheetVisible_(data, MB_SHEETS_QUALITY.docs);
  Logger.log('✅ Docs: ' + data.rows.length + ' rows → "' + MB_SHEETS_QUALITY.docs + '"');
}


// Master sync — called every 5 minutes by the installed trigger.
// Each card is wrapped independently so one failure doesn't block others.
function syncAllOnboarding() {
  Logger.log('═══ Metabase sync starting — ' + new Date().toISOString() + ' ═══');
  try { syncSellers(); } catch (e) { Logger.log('❌ Sellers error: ' + e.message); }
  try { syncBuyers();  } catch (e) { Logger.log('❌ Buyers error: '  + e.message); }
  try { syncDetail();  } catch (e) { Logger.log('❌ Detail error: '  + e.message); }
  try { syncRating();  } catch (e) { Logger.log('❌ Rating error: '  + e.message); }
  try { syncOSV();     } catch (e) { Logger.log('❌ OSV error: '     + e.message); }
  try { syncDocs();    } catch (e) { Logger.log('❌ Docs error: '    + e.message); }
  bustDashboardCache_();
  // Pre-warm caches for every period the UI exposes so user requests are always cache hits.
  // Individual getDashboardData calls are faster and also populate the vrows_v16_ caches
  // used by getVerticalRows; getCombinedDashboard composes from those when available.
  var warmPeriods = ['All', 'Today', 'ThisWeek', 'ThisMonth', 'MTD', 'YTD'];
  warmPeriods.forEach(function(p) {
    try {
      getDashboardData(JSON.stringify({ audience: 'seller', period: p }));
      getDashboardData(JSON.stringify({ audience: 'buyer',  period: p }));
    } catch (e) { Logger.log('  Pre-warm seller/buyer ' + p + ' failed: ' + e.message); }
    try {
      getCombinedDashboard(JSON.stringify({ period: p }));
    } catch (e) { Logger.log('  Pre-warm combined ' + p + ' failed: ' + e.message); }
  });
  Logger.log('  Cache pre-warmed (' + warmPeriods.length + ' periods × 3 calls each).');
  Logger.log('═══ Metabase sync complete — ' + new Date().toISOString() + ' ═══');
}


// ════════════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ════════════════════════════════════════════════════════════════

// Run ONCE to install the 5-minute auto-sync trigger.
function setupMetabaseTrigger() {
  clearMetabaseTrigger();   // remove any duplicate triggers first
  ScriptApp.newTrigger('syncAllOnboarding')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('✅ 5-minute trigger installed: syncAllOnboarding()');
}

// Run to stop auto-sync (e.g. for debugging or maintenance).
function clearMetabaseTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncAllOnboarding') {
      ScriptApp.deleteTrigger(t);
      Logger.log('  Removed trigger: ' + t.getUniqueId());
    }
  });
}


// ════════════════════════════════════════════════════════════════
// DIAGNOSTICS — run from the Apps Script editor to verify setup
// ════════════════════════════════════════════════════════════════
function testMetabaseConnection() {
  Logger.log('── Connection test ──');
  try {
    var token = getMBToken_();
    Logger.log('✅ Auth OK  (token prefix: ' + token.slice(0, 8) + '…)');
  } catch (e) {
    Logger.log('❌ Auth FAILED: ' + e.message);
    return;
  }

  [
    { label: 'Sellers',            id: MB_CARDS.seller         },
    { label: 'Buyers',             id: MB_CARDS.buyer          },
    { label: 'Onboarding Detail',  id: MB_CARDS.detail         },
    { label: 'Vendor Rating',      id: MB_CARDS_QUALITY.rating },
    { label: 'OSV Status',         id: MB_CARDS_QUALITY.osv    },
    { label: 'Doc Completeness',   id: MB_CARDS_QUALITY.docs   },
  ].forEach(function(card) {
    try {
      var d = fetchMBCard_(card.id);
      Logger.log('✅ Card ' + card.id + ' (' + card.label + '): ' +
        d.rows.length + ' rows | columns: ' + d.headers.join(', '));
    } catch (e) {
      Logger.log('❌ Card ' + card.id + ' (' + card.label + '): ' + e.message);
    }
  });

  Logger.log('── Sync timestamps ──');
  var props = PropertiesService.getScriptProperties();
  Object.values(MB_SHEETS).forEach(function(name) {
    var ts = props.getProperty('MB_SYNC_' + name);
    Logger.log('  ' + name + ': ' + (ts || 'never synced'));
  });
}

// Returns the last-sync timestamp for a given local sheet name.
// Called by Code.gs to annotate the "Last updated" timestamp in
// the dashboard response with the Metabase sync time when active.
function getMBSyncTime(sheetName) {
  return PropertiesService.getScriptProperties()
    .getProperty('MB_SYNC_' + sheetName) || null;
}
