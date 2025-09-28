/******************************************************
 * Ins Sheet setup + dropdown on A2:A (from "Function"!A2:A)
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/

function setInstSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Ins Sheet(DNT!)';

  // Labels & formulas
  const SELECT_FN_LABEL = 'Select Sheet Function =>';
  const LINK_OK_FORMULA = '=IF(REGEXMATCH(C2,"Form"),"Link Successful",IF(ISBLANK(E2),,IMPORTRANGE(C2,""&E2&"!B1")))';
  const INS_SHEET_FORM  = '=ARRAYFORMULA(IF(ISBLANK(C2:C),,IF(ISNUMBER(SEARCH("Corporate Application",C2:C,1)),,IF(REGEXMATCH(C2:C,"Form"),"https://docs.google.com/forms/d/"&RIGHT(C2:C.url,LEN(C2:C.url)-FIND("=",C2:C.url,1)),"Ins Sheet(DNT!)"))))';
  const URL_FORM        = '=ARRAYFORMULA(IF(ISBLANK(C2:C),,C2:C.url))';

  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);

    // Headers
    sh.getRange('A1').setValue(SELECT_FN_LABEL);
    sh.getRange('B1').setValue('Link Successful');
    sh.getRange('C1').setValue('Link here');
    sh.getRange('D1').setValue('Url List');
    sh.getRange('E1').setValue('Data List');

    // Header formatting
    setCellFormatting(sh, 'A1', 'white', 'bold', 'darkblue');
    setCellFormatting(sh, 'B1', 'white', 'bold', '#34a853');
    setCellFormatting(sh, 'C1', 'white', 'bold', '#00008b');
    setCellFormatting(sh, 'D1', 'white', 'bold', '#34a853');
    setCellFormatting(sh, 'E1', 'white', 'bold', '#34a853');

    // Formulas
    sh.getRange('B2:B').setFormula(LINK_OK_FORMULA); // status
    sh.getRange('D2').setFormula(URL_FORM);          // URL extract
    sh.getRange('E2').setFormula(INS_SHEET_FORM);    // data list helper

    // Conditional formatting on B2:B
    applyConditionalFormatting(sh, sh.getRange('B2:B'), '=B2=$B$1', '#34a853');      // green when matches "Link Successful"
    applyConditionalFormatting(sh, sh.getRange('B2:B'), '=ISBLANK(D2)=FALSE', 'red'); // red when D has value
  }

  // ✅ Build/refresh dropdown on A2:A from "Function"!A2:A
  buildInsRoleDropdown_();
}

/** Simple formatter */
function setCellFormatting(sheet, a1, fontColor, fontWeight, backgroundColor) {
  const r = sheet.getRange(a1);
  r.setFontColor(fontColor).setFontWeight(fontWeight).setBackground(backgroundColor);
}

/** Add a CF rule to a range */
function applyConditionalFormatting(sheet, range, customFormula, backgroundColor) {
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(customFormula)
    .setBackground(backgroundColor)
    .setFontColor('white')
    .setBold(true)
    .setRanges([range])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

/** Build dropdown on Ins Sheet A2:A using unique non-empty values from Function!A2:A */
function buildInsRoleDropdown_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('Function');
  const target = ss.getSheetByName('Ins Sheet(DNT!)');
  if (!source || !target) return; // quietly skip if either sheet missing

  const last = source.getLastRow();
  if (last < 2) return; // nothing to use

  const values2d = source.getRange(2, 1, last - 1, 1).getValues(); // A2:A
  const choices = Array.from(new Set(
    values2d.map(r => String(r[0] || '').trim()).filter(Boolean)
  ));
  if (!choices.length) return;

  const range = target.getRange('A2:A'); // whole column A from row 2
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(choices, true) // show dropdown
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}
