/** Custom menu and setup actions. */

function onOpen() {
  try { resetIfCopied_(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { /* not yet authorised */ }
  SpreadsheetApp.getUi().createMenu('Coding')
    .addItem('Set watched folder…', 'menuSetFolder')
    .addItem('Sync now', 'menuSyncNow')
    .addItem('Full re-sync (re-read every comment)', 'menuFullResync')
    .addSeparator()
    .addItem('Start auto-sync (every 10 min)', 'menuStartAuto')
    .addItem('Stop auto-sync', 'menuStopAuto')
    .addSeparator()
    .addItem('FigJam connection details', 'menuWebDetails')
    .addItem('Rotate FigJam token', 'menuRotateToken')
    .addItem('Reset web app URL', 'menuResetWebUrl')
    .addItem('Unlink FigJam board', 'menuUnlinkBoard')
    .addSeparator()
    .addItem('Run self-tests', 'runTests')
    .addToUi();
}

function menuSetFolder() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Watched folder', 'Paste the Google Drive folder URL (or ID):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var text = res.getResponseText().trim();
  var id = (text.match(/folders\/([\w-]+)/) || [])[1] || text;
  var folder;
  try { folder = DriveApp.getFolderById(id); } catch (e) {
    ui.alert('Could not open that folder. Check the link and that you have access.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('FOLDER_ID', id);
  ensureSheets_();
  ui.alert('Watching "' + folder.getName() + '" (and its subfolders). Run Coding → Sync now.');
}

function menuSyncNow() {
  var s = syncAll();
  if (s && s.skipped) {
    SpreadsheetApp.getUi().alert('A sync is already running. Try again in a minute.');
    return;
  }
  SpreadsheetApp.getActive().toast(
    s.added + ' added, ' + s.updated + ' updated, ' + s.deleted + ' marked deleted. See the Log tab.',
    'Sync done', 8);
}

/** Forget per-file sync times so every comment is re-read. Rows are updated in place, not duplicated. */
function menuFullResync() {
  ensureSheets_();
  var sh = ss_().getSheetByName(SHEETS.SOURCES);
  var n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, 5, n, 1).clearContent();
    sh.getRange(2, 7, n, 1).clearContent();
  }
  menuSyncNow();
}

// Time triggers belong to the person who creates them and are invisible to
// everyone else, so record who runs auto-sync: it only needs one person.
function menuStartAuto() {
  var props = PropertiesService.getScriptProperties();
  var me = Session.getEffectiveUser().getEmail();
  var owner = props.getProperty('AUTO_SYNC_BY');
  if (owner && owner !== me) {
    SpreadsheetApp.getUi().alert('Auto-sync is already on, run by ' + owner + '. One person is enough.');
    return;
  }
  deleteMyTriggers_();
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(10).create();
  props.setProperty('AUTO_SYNC_BY', me);
  SpreadsheetApp.getActive().toast('Auto-sync on: every 10 minutes, run by you.', 'Coding', 5);
}

function menuStopAuto() {
  var props = PropertiesService.getScriptProperties();
  var me = Session.getEffectiveUser().getEmail();
  var owner = props.getProperty('AUTO_SYNC_BY');
  if (owner && owner !== me) {
    SpreadsheetApp.getUi().alert('Auto-sync is run by ' + owner + '. Ask them to stop it from this menu.');
    return;
  }
  deleteMyTriggers_();
  props.deleteProperty('AUTO_SYNC_BY');
  SpreadsheetApp.getActive().toast('Auto-sync off.', 'Coding', 5);
}

function deleteMyTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
}

function menuUnlinkBoard() {
  var props = PropertiesService.getScriptProperties();
  var name = props.getProperty('BOARD_NAME');
  props.deleteProperty('BOARD_ID');
  props.deleteProperty('BOARD_NAME');
  SpreadsheetApp.getUi().alert(name
    ? 'Unlinked "' + name + '". The next board to Sync becomes the linked board.'
    : 'No board was linked.');
}

function menuWebDetails() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('WEB_TOKEN')) props.setProperty('WEB_TOKEN', Utilities.getUuid());
  ensureSheets_();
  // ScriptApp.getService().getUrl() returns a dead URL on some Workspace domains,
  // so ask once for the real one and remember it.
  var url = props.getProperty('WEB_URL');
  if (!url) {
    var res = ui.prompt('Web app URL',
      'In the Apps Script editor: Deploy → Manage deployments → copy the Web app URL ' +
      '(ends in /exec) and paste it here.\n\nNot deployed yet? Deploy → New deployment → Web app, ' +
      'Execute as Me, Who has access Anyone.', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    url = normaliseWebUrl_(res.getResponseText());
    if (!url) { ui.alert('That doesn\'t look like a web app URL ending in /exec.'); return; }
    props.setProperty('WEB_URL', url);
  }
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px sans-serif;line-height:1.5">' +
    '<p><b>Web app URL</b><br><input style="width:100%" readonly value="' + url + '"></p>' +
    '<p><b>Token</b><br><input style="width:100%" readonly value="' + props.getProperty('WEB_TOKEN') + '"></p>' +
    '<p><b>Linked board</b><br>' + (props.getProperty('BOARD_NAME') || '(none yet: the first board to Sync is linked)') + '</p>' +
    '<p style="color:#555">Paste both into the FigJam plugin. Anyone with both can read the coded ' +
    'excerpts, so keep them to yourself. <i>Rotate FigJam token</i> issues a new one; ' +
    '<i>Reset web app URL</i> if you redeploy under a new address.</p></div>'
  ).setWidth(460).setHeight(360);
  ui.showModalDialog(html, 'FigJam connection');
}

/** Accepts /macros/s/ID/exec, /a/macros/DOMAIN/s/ID/exec or /a/DOMAIN/macros/s/ID/…; returns the plain /exec form. */
function normaliseWebUrl_(text) {
  var id = (String(text).match(/\/s\/([\w-]{20,})/) || [])[1];
  return id ? 'https://script.google.com/macros/s/' + id + '/exec' : '';
}

function menuResetWebUrl() {
  PropertiesService.getScriptProperties().deleteProperty('WEB_URL');
  menuWebDetails();
}

function menuRotateToken() {
  PropertiesService.getScriptProperties().setProperty('WEB_TOKEN', Utilities.getUuid());
  menuWebDetails();
}
