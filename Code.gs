// ═══════════════════════════════════════════════════════════════
// ENTERPRISE BUSINESS CONTROL TOWER — Apps Script entry point
// ═══════════════════════════════════════════════════════════════
//
// The dashboard now reads its data DIRECTLY IN THE BROWSER from two Google
// Sheets (Sellers + Buyers) via the Google Visualization endpoint — see the
// DATA LAYER section in Index.html (SHEETS config + fetchSheetRows).
//
// There is no server-side sheet processing here anymore; the previous
// Marketplace / Open Marketplace / EPR sheet connections have been removed.
// This file only serves the single-page app when deployed as an Apps Script
// web app. (If Index.html is hosted statically instead, this file is unused.)
//
// Sheets consumed by the frontend:
//   Sellers : 1qaG_GMvUrC7LJbKBma8-S3x2N6Ua4vDkUC5ZRx3znho
//   Buyers  : 1G9Ocq8PXovCx5eBE3dOfE8CUcmfgwcl6Te37p79u0wI
// Both must be shared "Anyone with the link → Viewer" so the browser can read
// them without authentication.

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enterprise Business Control Tower')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
