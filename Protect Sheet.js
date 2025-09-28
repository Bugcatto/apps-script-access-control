/******************************************************
 * Sheet protection for "Ins Sheet(DNT!)" (protection.gs)
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 *
 * Depends on config.gs:
 *  - MAINTENANCE_EMAILS
 *  - openHostSpreadsheet_()
 ******************************************************/

/** Back-compat export; keep your old name */
function protectSheetinstasheet() { protectInsSheet(); }

/** Main: protect the Ins sheet and allow only Admin + backdoor + owner */
function protectInsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Ins Sheet(DNT!)');
  if (!sh) throw new Error('Sheet "Ins Sheet(DNT!)" not found.');

  // Build allowed editors
  const owner = ss.getOwner().getEmail().toLowerCase();
  const admins = new Set(getSystemAdministratorEmails_().map(e => e.toLowerCase()));
  MAINTENANCE_EMAILS.forEach(e => admins.add(String(e).toLowerCase()));
  admins.add(owner); // always keep owner

  // Reuse existing protection if any, else create once
  let prot = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).find(p => true);
  if (!prot) {
    prot = sh.protect();
    prot.setDescription('Sheet Protection: Ins Sheet(DNT!)');
  }

  prot.setWarningOnly(false); // enforce, not just warn

  // Reset editors (clear then add allowed)
  // Note: Protection editors are independent from file editors.
  prot.getEditors().forEach(u => {
    const em = u.getEmail().toLowerCase();
    if (!admins.has(em)) {
      try { prot.removeEditor(em); } catch (_) {}
    }
  });
  admins.forEach(em => {
    try { prot.addEditor(em); } catch (_) {}
  });

  // Optional: If you want only certain ranges editable, uncomment below:
  // prot.setUnprotectedRanges([ sh.getRange('A1') ]); // example

  ss.toast('Ins Sheet protected; editors limited to Admin/Backdoor/Owner.', 'Protection', 4);
}

/** Admin list = "System Administrator" column from host "Access Control" (B1:S*) */
function getSystemAdministratorEmails_() {
  const host = openHostSpreadsheet_();
  const sheet = host.getSheetByName('Access Control');
  if (!sheet) throw new Error('Host sheet "Access Control" not found.');

  const rng = sheet.getRange(1, 2, sheet.getLastRow(), 18); // B1:S*
  const data = rng.getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h || '').trim());
  const colIdx = headers.indexOf('System Administrator');
  if (colIdx === -1) return [];

  const emails = [];
  for (let r = 1; r < data.length; r++) {
    const em = String(data[r][colIdx] || '').trim();
    if (em) emails.push(em.toLowerCase());
  }
  return Array.from(new Set(emails));
}

/** Thin alias if you want to keep this callable elsewhere */
function deleteAllTriggers() { deleteAllMyProjectTriggers_(); }
