/******************************************************
 * Folder Permission Hardening (driveHarden.gs)
 * Author: Phone Myat
 * Contact: phonemyatthanoo@gmail.com, +9595350911
 *
 * Depends on config.gs:
 *  - getMainFolderId_()
 *  - MAINTENANCE_EMAILS
 *  - ALLOWLIST_EDITORS (optional)
 *  - ALLOWLIST_VIEWERS (optional)
 ******************************************************/

/** Entrypoint */
function setPermissionsForFolder() {
  const folderId = getMainFolderId_(); // throws if not set
  const folder = DriveApp.getFolderById(folderId);
  processFolder_(folder);
  SpreadsheetApp.getActive().toast('Folder permissions hardened.', 'Drive Security', 4);
}

/** DFS over folder tree */
function processFolder_(folder) {
  // Files
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      setPermissionsForFile_(file);
      Utilities.sleep(50); // be gentle with quotas
    } catch (e) {
      Logger.log('File perm error ' + file.getId() + ': ' + e);
    }
  }
  // Subfolders
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    processFolder_(subs.next());
  }
}

/** Harden a single file */
function setPermissionsForFile_(file) {
  const fileId = file.getId();

  // 1) Make private (removes public link)
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

  // 2) Ensure maintenance editors present
  const currentEditors = new Set(file.getEditors().map(u => u.getEmail().toLowerCase()));
  (MAINTENANCE_EMAILS || []).forEach(e => {
    const em = String(e || '').toLowerCase().trim();
    if (em && !currentEditors.has(em)) {
      try { file.addEditor(em); Logger.log('Ensured maint editor: ' + em); }
      catch (err) { Logger.log('Cannot add maint editor ' + em + ': ' + err); }
    }
  });

  // 3) (Optional) prune editors/viewers not in allowlists
  const allowEditors = (typeof ALLOWLIST_EDITORS !== 'undefined' && ALLOWLIST_EDITORS instanceof Set)
    ? ALLOWLIST_EDITORS
    : new Set((MAINTENANCE_EMAILS || []).map(e => String(e).toLowerCase()));
  const allowViewers = (typeof ALLOWLIST_VIEWERS !== 'undefined' && ALLOWLIST_VIEWERS instanceof Set)
    ? ALLOWLIST_VIEWERS
    : new Set(); // default: remove all viewers

  if (allowEditors.size > 0) {
    file.getEditors().forEach(u => {
      const em = u.getEmail().toLowerCase();
      if (!allowEditors.has(em)) {
        try { file.removeEditor(em); Logger.log('Removed editor: ' + em); }
        catch (err) { Logger.log('Skip remove editor ' + em + ': ' + err); }
      }
    });
  }
  if (allowViewers.size >= 0) { // empty set => remove all viewers
    file.getViewers().forEach(u => {
      const em = u.getEmail().toLowerCase();
      if (!allowViewers.has(em)) {
        try { file.removeViewer(em); Logger.log('Removed viewer: ' + em); }
        catch (err) { Logger.log('Skip remove viewer ' + em + ': ' + err); }
      }
    });
  }

  // 4) Block resharing & copying
  const resource = {
    copyRequiresWriterPermission: true,
    viewersCanCopyContent: false,
    writersCanShare: false
  };

  try {
    // Advanced Drive Service (faster) if enabled
    if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.update) {
      Drive.Files.update(resource, fileId, null, { supportsAllDrives: true });
    } else {
      // REST fallback
      const url = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?supportsAllDrives=true';
      const params = {
        method: 'patch',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify(resource),
        muteHttpExceptions: true
      };
      const resp = UrlFetchApp.fetch(url, params);
      if (resp.getResponseCode() >= 300) {
        Logger.log('Drive v3 patch failed: ' + resp.getContentText());
      }
    }
  } catch (e) {
    Logger.log('Drive API update failed for ' + fileId + ': ' + e);
  }
}
