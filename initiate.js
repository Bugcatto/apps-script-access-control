/******************************************************
 * Main Project — Setup & Triggers (no library calls)
 * Author: Phone Myat | Contact: phonemyatthanoo@gmail.com, +9595350911
 ******************************************************/

//const M_TITLE = 'Function';

/** One-time setup */
function Initialsetup() {
  // Ensure required sheets / config (your local creator)
  safeCall_(setInstSheet);
  safeCall_(setAccess);
  safeCall_(setempinfo);

  // Fresh triggers (avoid duplicates)
  deleteAllMyProjectTriggers_();
  createOnOpenTrigger_();
  createHourlySetAccessTrigger_();
  createDailyBackupTrigger_();
  createWeeklyBackupTrigger_();

  SpreadsheetApp.getActive().toast('Setup complete.', 'Setup', 5);
}

/** onOpen handler */
function onOpenTrigger() {
  // Keep access fresh & show directory UI
  safeCall_(setAccess);
  safeCall_(showUrls);

  // Menu
  SpreadsheetApp.getUi()
    .createMenu(M_TITLE)
    .addItem('Update Access', 'setAccess')
    .addItem('Directory', 'showUrls')
    .addSeparator()
    .addItem('Repair Triggers', 'repairTriggers')
    .addSeparator()
    .addItem('Byebye (Kick All Except Owner)', 'byebye')
    .addToUi();
}

/** ===== Public maint action ===== */
function repairTriggers() {
  dedupeTriggers_(['onOpenTrigger','SetAccess','callbackup24','callbackup7']);
  if (!existsTrigger_('onOpenTrigger'))        createOnOpenTrigger_();
  if (!existsTrigger_('SetAccess'))            createHourlySetAccessTrigger_();
  if (!existsTrigger_('callbackup24'))         createDailyBackupTrigger_();
  if (!existsTrigger_('callbackup7'))          createWeeklyBackupTrigger_();
  SpreadsheetApp.getActive().toast('Triggers repaired.', 'Triggers', 4);
}

/** ===== Trigger creators (idempotent helpers) ===== */
function createOnOpenTrigger_() {
  deleteTriggers_(['onOpenTrigger']);
  ScriptApp.newTrigger('onOpenTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

function createHourlySetAccessTrigger_() {
  deleteTriggers_(['SetAccess']);
  ScriptApp.newTrigger('SetAccess').timeBased().everyHours(1).create();
}

function createDailyBackupTrigger_() {
  deleteTriggers_(['callbackup24']);
  ScriptApp.newTrigger('callbackup24').timeBased().everyDays(1).atHour(0).create();
}

function createWeeklyBackupTrigger_() {
  deleteTriggers_(['callbackup7']);
  ScriptApp.newTrigger('callbackup7').timeBased().everyDays(7).atHour(0).create();
}

/** Backup callbacks (call your local implementations) */
function callbackup24() { safeCall_(backupSpreadsheet24); }
function callbackup7()  { safeCall_(backupSpreadsheet7);  }

/** ===== Utilities ===== */
function deleteAllMyProjectTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    try { ScriptApp.deleteTrigger(t); } catch (e) { Logger.log('Delete skip: ' + e); }
  });
}
function deleteTriggers_(handlers) {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) {
      try { ScriptApp.deleteTrigger(t); } catch (e) { Logger.log('Delete skip: ' + e); }
    }
  });
}
function existsTrigger_(handler) {
  return ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction() === handler; });
}
function dedupeTriggers_(handlers) {
  var all = ScriptApp.getProjectTriggers();
  handlers.forEach(function(h){
    var list = all.filter(function(t){ return t.getHandlerFunction() === h; });
    list.slice(1).forEach(function(t){ try { ScriptApp.deleteTrigger(t); } catch(e){} });
  });
}
function safeCall_(fn){ try { if (typeof fn === 'function') fn(); } catch(e){ Logger.log(e); } }
