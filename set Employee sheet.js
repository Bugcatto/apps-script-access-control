/******************************************************
 * Employee sheet setup (empInfo.gs)
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/
function setempinfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const SHEET_NAME = 'Employee Info';

    // Create or clear (keep formatting if sheet exists)
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
    } else {
      sh.clearContents().clearNotes();
    }

    // Seed headers (adjust as you like)
    const headers = [
      'emp_id','name','email','phone','role','department',
      'status','joined_on','last_updated'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);

    // Optional: size a few columns
    sh.autoResizeColumns(1, headers.length);

    // If you have a data-fetcher, call it with the host URL
    if (typeof getQueryData === 'function') {
      getQueryData(getHostSpreadsheetUrl_()); // from config.gs
    }

    ss.toast('Employee Info sheet ready.', 'Setup', 4);
  } catch (err) {
    Logger.log('setempinfo error: ' + err);
    SpreadsheetApp.getActive().toast('Employee setup error. Check logs.', 'Error', 6);
  }
}
