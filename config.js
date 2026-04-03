/******************************************************
 * CONFIG & PROPERTIES (config.gs)
 * Bookstore Admin
 * Author: Phone Myat
 * Contact: phonemyatthanoo@gmail.com, +9595350911
 *
 * Purpose:
 * - Single source of truth for constants & Script Properties
 * - Helpers to safely read properties (no stale cached globals)
 * - Minimal UI menu (logic lives in other files)
 ******************************************************/

/** ===== Constants ===== */
const M_TITLE = 'Function';                                 // Menu title
const PROP    = PropertiesService.getScriptProperties();    // Script properties handle

// Property keys (no spaces)
const K_SPREADSHEET_URL  = 'spreadsheetUrl';     // Host/remote SS (e.g., Access Control)
const K_DSPREADSHEET_URL = 'dspreadsheetUrl';    // Optional secondary/directory SS
const K_FOLDER_ID        = 'folder_id';          // Drive root folder to secure
const K_DIRECTORY_SHEET  = 'directory_sheet';    // Directory sheet name inside host/dir SS

// Maintenance backdoor (always keep editor access)
const MAINTENANCE_EMAILS = ['phonemyatthanoo@gmail.com','bookscanner@bookstore-inventory-ocr.iam.gserviceaccount.com'];

// Optional allowlists used by Drive hardening (other files can import)
const ALLOWLIST_EDITORS = new Set(MAINTENANCE_EMAILS.map(e => e.toLowerCase()));
const ALLOWLIST_VIEWERS = new Set(); // keep empty to remove all viewers

/** ===== One-time initializer (run manually once) ===== */
function setSpreadsheetUrlProperty() {
  const url      = 'https://docs.google.com/spreadsheets/d/1porRSTkyW4qHeKayyd-8wnuBft6-HFszeJMc4pBW1Is/edit';
  const durl     = 'https://docs.google.com/spreadsheets/d/1FERzh4tiYgfRYQLCet9FLZQ5ztNHQo6tQ1qcbPFFsFo/edit';
  const folderId = '1yG-rJEE3n9A6idVWRcl4Ps8iYoTtfCS-';
  const dirSheet = 'Sheet Directory';

  PROP.setProperty(K_SPREADSHEET_URL,  url);
  PROP.setProperty(K_DSPREADSHEET_URL, durl);
  PROP.setProperty(K_FOLDER_ID,        folderId);
  PROP.setProperty(K_DIRECTORY_SHEET,  dirSheet);

  SpreadsheetApp.getActive().toast('Script properties initialized.', 'Setup', 5);
}

/** ===== Safe property accessors (no stale globals) ===== */
function getProp_(key) { return (PROP.getProperty(key) || '').trim(); }

function getHostSpreadsheetUrl_() {
  const v = getProp_(K_SPREADSHEET_URL);
  if (!v) throw new Error('spreadsheetUrl not set. Run setSpreadsheetUrlProperty().');
  return v;
}
function getDirectorySpreadsheetUrl_() { return getProp_(K_DSPREADSHEET_URL); }
function getMainFolderId_() {
  const v = getProp_(K_FOLDER_ID);
  if (!v) throw new Error('folder_id not set. Run setSpreadsheetUrlProperty().');
  return v;
}
function getDirectorySheetName_() {
  const v = getProp_(K_DIRECTORY_SHEET);
  return v || 'Sheet Directory';
}

/** Open host spreadsheet by URL property (used by import/access scripts) */
function openHostSpreadsheet_() {
  return SpreadsheetApp.openByUrl(getHostSpreadsheetUrl_());
}

/** ===== Minimal menu (keep logic in other files) ===== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(M_TITLE)
    .addItem('Update Access', 'setAccess') // implemented in access.gs
    .addToUi();
}
