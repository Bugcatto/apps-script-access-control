/******************************************************
 * Import emails → "View Permission List"
 * Sync access on destination SS + Drive folder
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/
function importDataAndViewAccess() {
  try {
    // --- Open sources via config helpers ---
    const hostSS = openHostSpreadsheet_();                          // from config.gs
    const destUrl = getDirectorySpreadsheetUrl_();                  // may throw if not set
    if (!destUrl) throw new Error('Directory spreadsheet URL not set.');
    const destSS  = SpreadsheetApp.openByUrl(destUrl);

    const folderId = getMainFolderId_();                            // from config.gs
    const folder   = DriveApp.getFolderById(folderId);

    // --- Sheets ---
    const srcSheet  = hostSS.getSheetByName('Folder Accsess Control');
    if (!srcSheet) throw new Error(`Source sheet "Folder Accsess Control" not found.`);
    const destSheet = destSS.getSheetByName('View Permission List');
    if (!destSheet) throw new Error(`Destination sheet "View Permission List" not found.`);

    // --- Read + clean emails from A2:A (ignore blanks/errors) ---
    const srcLast = srcSheet.getLastRow();
    if (srcLast < 2) throw new Error('No source emails to import.');
    const raw2D   = srcSheet.getRange(2, 1, srcLast - 1, 1).getValues();

    const emails = Array.from(new Set(
      raw2D
        .map(r => String(r[0] || '').trim())
        .filter(v => v && !/#(REF|VALUE|DIV\/0|NAME|N\/A|NUM|NULL)!/i.test(v))
        .map(v => v.toLowerCase())
    ));

    // --- Write cleaned list to destination A2:A (overwrite old) ---
    const oldLast = destSheet.getLastRow();
    if (oldLast > 1) destSheet.getRange(2, 1, oldLast - 1, 1).clearContent();
    if (emails.length) {
      const out = emails.map(e => [e]);
      destSheet.getRange(2, 1, out.length, 1).setValues(out);
    }

    // --- Preserve owner + maintenance backdoor ---
    const ownerEmail = destSS.getOwner().getEmail().toLowerCase();
    const mustKeep   = new Set([ownerEmail, ...MAINTENANCE_EMAILS.map(e => String(e).toLowerCase())]);
    mustKeep.forEach(e => emails.includes(e) || emails.push(e));

    // --- Current permissions -> Sets for fast diff ---
    const viewersSet = new Set(destSS.getViewers().map(u => u.getEmail().toLowerCase()));
    const editorsSet = new Set(folder.getEditors().map(u => u.getEmail().toLowerCase()));
    const wantSet    = new Set(emails);

    // --- Grant missing (viewer on SS, editor on folder) ---
    emails.forEach(e => {
      if (!viewersSet.has(e)) { try { destSS.addViewer(e); } catch (err) { Logger.log('addViewer fail ' + e + ': ' + err); } }
      if (!editorsSet.has(e)) { try { folder.addEditor(e); } catch (err) { Logger.log('addEditor fail ' + e + ': ' + err); } }
    });

    // --- Revoke extras (not in wantSet), but NEVER remove mustKeep ---
    destSS.getViewers().forEach(u => {
      const e = u.getEmail().toLowerCase();
      if (!mustKeep.has(e) && !wantSet.has(e)) {
        try { destSS.removeViewer(e); } catch (err) { Logger.log('removeViewer fail ' + e + ': ' + err); }
      }
    });
    folder.getEditors().forEach(u => {
      const e = u.getEmail().toLowerCase();
      if (!mustKeep.has(e) && !wantSet.has(e)) {
        try { folder.removeEditor(e); } catch (err) { Logger.log('removeEditor fail ' + e + ': ' + err); }
      }
    });

    SpreadsheetApp.getActive().toast('Directory access synced.', 'Success', 4);
  } catch (err) {
    Logger.log('importDataAndViewAccess error: ' + err);
    SpreadsheetApp.getUi().alert('Error: ' + err);
  }
}

/** (Optional) simple email validator if you want extra filtering
function isValidEmail_(e){ return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(e||''); }
*/
