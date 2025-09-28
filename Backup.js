/******************************************************
 * Unified Spreadsheet Backup (daily/weekly) — backup.gs
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/

function callbackup24() { backupDaily(); }   // existing daily trigger target
function callbackup7()  { backupWeekly(); }  // existing weekly trigger target

function backupDaily()  { backupSpreadsheet_({ frequency: 'daily',  keepLast: 7  }); }
function backupWeekly() { backupSpreadsheet_({ frequency: 'weekly', keepLast: 5  }); }

/**
 * Create a timestamped copy of the active spreadsheet under:
 *   <Backup Root>/Backup/Daily  or  <Backup Root>/Backup/Weekly
 * Then delete older backups beyond keepLast.
 *
 * @param {{frequency:'daily'|'weekly', keepLast?:number}} opts
 */
function backupSpreadsheet_(opts) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30 * 1000);

  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const tz         = 'Asia/Yangon';
    const stamp      = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmm');
    const freq       = (opts && opts.frequency) || 'daily';
    const keepLast   = (opts && opts.keepLast) || (freq === 'weekly' ? 5 : 7);
    const activeName = ss.getName();
    const backupName = `[${freq.toUpperCase()}] ${activeName} - ${stamp}`;

    // Choose backup root folder:
    // If you have getMainFolderId_() in config.gs, use it; else fallback to Drive root.
    const rootFolder = (typeof getMainFolderId_ === 'function')
      ? DriveApp.getFolderById(getMainFolderId_())
      : DriveApp.getRootFolder();

    // Ensure /Backup/<Daily|Weekly> path
    const backupRoot = ensureSubfolder_(rootFolder, 'Backup');
    const freqFolder = ensureSubfolder_(backupRoot, (freq === 'weekly') ? 'Weekly' : 'Daily');

    // Full file copy (fast, reliable)
    const srcFile  = DriveApp.getFileById(ss.getId());
    const copyFile = srcFile.makeCopy(backupName, freqFolder);

    // Retention
    enforceRetention_(freqFolder, keepLast);

    ss.toast(`Backup created: ${copyFile.getName()} (kept last ${keepLast})`, 'Backup', 5);
  } catch (e) {
    Logger.log('backupSpreadsheet_ error: ' + e);
    SpreadsheetApp.getActive().toast('Backup error. Check logs.', 'Error', 6);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Ensure a subfolder exists by name under parent; return it */
function ensureSubfolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

/** Keep only the most recent N files in a folder (by created date desc) */
function enforceRetention_(folder, keepLast) {
  const files = [];
  for (let it = folder.getFiles(); it.hasNext(); ) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(keepLast).forEach(f => {
    try { f.setTrashed(true); Logger.log('Deleted old backup: ' + f.getName()); }
    catch (e) { Logger.log('Failed to delete ' + f.getName() + ': ' + e); }
  });
}
