/******************************************************
 * Import "Function List" → local "Function"
 * Build dropdown on "Ins Sheet(DNT!)" A2:A
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/

function importfunction(sourceSheetName = 'Function List',
                        targetSheetName = 'Function',
                        buildDropdown = true) {
  // Open host via config helper (throws if not set)
  const host = openHostSpreadsheet_();

  // Read source
  const src = host.getSheetByName(sourceSheetName);
  if (!src) throw new Error(`Source sheet "${sourceSheetName}" not found.`);
  const values = src.getDataRange().getValues();
  if (!values.length || !values[0].length) throw new Error(`"${sourceSheetName}" is empty.`);

  // Prepare target
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dest = ss.getSheetByName(targetSheetName);
  dest ? dest.clearContents().clearFormats() : dest = ss.insertSheet(targetSheetName);

  // Write & tidy
  dest.getRange(1, 1, values.length, values[0].length).setValues(values);
  if (values.length > 1) dest.setFrozenRows(1);
  dest.autoResizeColumns(1, dest.getLastColumn());

  // Optional dropdown build
  if (buildDropdown) {
    createDropdownList({
      listSheetName: targetSheetName,
      listColumn: 1,
      listStartRow: 2,
      targetSheetName: 'Ins Sheet(DNT!)',
      targetColumn: 1,
      targetStartRow: 2
    });
  }
}

function createDropdownList(opts) {
  const {
    listSheetName, listColumn, listStartRow,
    targetSheetName, targetColumn, targetStartRow,
    targetRowCount
  } = opts;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Source list
  const listSheet = ss.getSheetByName(listSheetName);
  if (!listSheet) throw new Error(`List sheet "${listSheetName}" not found.`);
  const lastListRow = listSheet.getLastRow();
  if (lastListRow < listStartRow) throw new Error(`No values in "${listSheetName}" from row ${listStartRow}.`);

  const height = lastListRow - listStartRow + 1;
  const raw = listSheet.getRange(listStartRow, listColumn, height, 1).getValues();

  // Clean + dedupe
  const choices = Array.from(new Set(raw.map(r => String(r[0]).trim()).filter(Boolean)));
  if (!choices.length) throw new Error('No valid choices for dropdown.');

  // Target
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) throw new Error(`Target sheet "${targetSheetName}" not found.`);

  const rowsToApply = (typeof targetRowCount === 'number' && targetRowCount > 0)
    ? targetRowCount
    : Math.max(targetSheet.getLastRow(), targetStartRow) - targetStartRow + 1;

  const range = targetSheet.getRange(targetStartRow, targetColumn, rowsToApply, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(choices, true)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
}
