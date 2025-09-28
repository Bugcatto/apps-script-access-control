/******************************************************
 * Pull "Role Assignment" → filter by roles in Ins Sheet(DNT!)
 * Writes result into "Employee Info"
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 *
 * @param {string=} hostUrl  Optional. If omitted, uses openHostSpreadsheet_() from config.
 ******************************************************/
function getQueryData(hostUrl) {
  try {
    // --- Host spreadsheet (source) ---
    const hostSS = hostUrl ? SpreadsheetApp.openByUrl(hostUrl) : openHostSpreadsheet_();
    const roleSheet = hostSS.getSheetByName('Role Assignment');
    if (!roleSheet) throw new Error('Source sheet "Role Assignment" not found.');

    const lastRow = roleSheet.getLastRow();
    const lastCol = roleSheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) throw new Error('"Role Assignment" is empty.');

    // Read header + data rows only
    const header = roleSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const body   = (lastRow > 1) ? roleSheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

    // --- Active spreadsheet (destination & filter) ---
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const ins = activeSS.getSheetByName('Ins Sheet(DNT!)');
    if (!ins) throw new Error('Local sheet "Ins Sheet(DNT!)" not found.');
    const emp = activeSS.getSheetByName('Employee Info') || activeSS.insertSheet('Employee Info');

    // Collect roles from A2:A (clean + dedupe)
    const insLast = ins.getLastRow();
    const roles = (insLast >= 2)
      ? Array.from(new Set(ins.getRange(2, 1, insLast - 1, 1).getValues()
          .map(r => String(r[0] || '').trim())
          .filter(Boolean)))
      : [];

    if (roles.length === 0) {
      // No filters → just clear contents (keep formatting)
      emp.clearContents().clearNotes();
      emp.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
      emp.setFrozenRows(1);
      Logger.log('Employee Info cleared (no roles selected).');
      return;
    }

    // Filter rows where first column (role) matches any selected role
    const roleSet = new Set(roles.map(r => r.toLowerCase()));
    const filtered = body.filter(row => {
      const role = String(row[0] || '').toLowerCase();
      return role && roleSet.has(role);
    });

    // Write output (header + filtered rows)
    emp.clearContents().clearNotes();
    emp.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    if (filtered.length) {
      emp.getRange(2, 1, filtered.length, header.length).setValues(filtered);
    }
    emp.setFrozenRows(1);
    emp.autoResizeColumns(1, header.length);

    Logger.log(`Employee Info updated: ${filtered.length} rows (of ${body.length}).`);
  } catch (err) {
    Logger.log('getQueryData error: ' + err);
    SpreadsheetApp.getUi().alert('getQueryData error: ' + err);
  }
}
