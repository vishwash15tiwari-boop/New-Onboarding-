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

// ── Local sheet names (created hidden in this spreadsheet) ────────
var MB_SHEETS = {
  seller: '_mb_sellers',
  buyer:  '_mb_buyers',
  detail: '_mb_detail',
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
  PropertiesService.getScriptProperties().setProperties({
    'MB_EMAIL':    'vishwash.tiwari@recykal.com',
    'MB_PASSWORD': '',   // ← paste password here, run once, then blank this
  });
  Logger.log('✅ Credentials saved to Script Properties.');
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
// Retries once on 401 (token expired mid-cycle).
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

  var res = doFetch_(token);

  if (res.getResponseCode() === 401) {
    // Token expired mid-cycle — refresh once and retry
    clearMBToken_();
    token = getMBToken_();
    res   = doFetch_(token);
  }

  if (res.getResponseCode() !== 200) {
    throw new Error('Card ' + cardId + ' fetch failed (' + res.getResponseCode() + '): ' +
      res.getContentText().slice(0, 300));
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

  // Clear existing content (preserve formatting)
  var usedRows = sheet.getLastRow();
  var usedCols = Math.max(sheet.getLastColumn(), payload[0].length);
  if (usedRows > 0) sheet.getRange(1, 1, usedRows, usedCols).clearContent();

  // Write header + data
  sheet.getRange(1, 1, payload.length, payload[0].length).setValues(payload);

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
  var periods = ['All', 'Today', 'ThisWeek', 'ThisMonth'];

  // Add FY keys for the last 5 years
  var nowYear = new Date().getFullYear();
  for (var y = nowYear - 4; y <= nowYear; y++) {
    periods.push('FY' + String(y).slice(2) + '-' + String(y + 1).slice(2));
  }

  ['seller', 'buyer'].forEach(function(aud) {
    periods.forEach(function(p) {
      keys.push('dash_v13_' + aud + '_' + JSON.stringify([p, '', '']));
    });
  });

  keys.forEach(function(k) { try { cache.remove(k); } catch (e) {} });
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

function syncDetail() {
  Logger.log('▶ Syncing Onboarding Detail (card ' + MB_CARDS.detail + ')…');
  var data = fetchMBCard_(MB_CARDS.detail);
  writeMBSheet_(data, MB_SHEETS.detail);
  Logger.log('✅ Detail: ' + data.rows.length + ' rows → "' + MB_SHEETS.detail + '"');
}

// Master sync — called every 5 minutes by the installed trigger.
// Each card is wrapped independently so one failure doesn't block others.
function syncAllOnboarding() {
  Logger.log('═══ Metabase sync starting — ' + new Date().toISOString() + ' ═══');
  try { syncSellers(); } catch (e) { Logger.log('❌ Sellers error: ' + e.message); }
  try { syncBuyers();  } catch (e) { Logger.log('❌ Buyers error: '  + e.message); }
  try { syncDetail();  } catch (e) { Logger.log('❌ Detail error: '  + e.message); }
  bustDashboardCache_();
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
    { label: 'Sellers',            id: MB_CARDS.seller },
    { label: 'Buyers',             id: MB_CARDS.buyer  },
    { label: 'Onboarding Detail',  id: MB_CARDS.detail },
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
