/**
 * Live feed for the FigJam plugin.
 *   GET  ?token=…                → {rows:[…]} (all non-deleted Codings rows)
 *   POST {token, stickies:[…]}   → recodes, board state, themes (see Board.gs)
 * POST bodies are sent as text/plain so browsers skip the CORS preflight.
 */

function doGet(e) {
  if (!tokenOk_(e.parameter.token)) return json_({ error: 'bad token' });
  var sh = ss_().getSheetByName(SHEETS.CODINGS);
  var n = sh.getLastRow() - 1;
  if (n < 1) return json_({ rows: [] });
  var width = SCRIPT_COLS + 3; // + participant, code, theme
  var values = sh.getRange(2, 1, n, width).getValues();
  var links = sh.getRange(2, C.NAME + 1, n, 1).getFormulas();
  var colours = codeColours_();
  var rows = [];
  values.forEach(function (r, i) {
    if (!r[C.KEY] || r[C.DELETED] === true) return;
    var link = (links[i][0].match(/^=HYPERLINK\("((?:[^"]|"")*)"/) || [])[1];
    var code = String(r[SCRIPT_COLS + 1]);
    rows.push({
      row_key: String(r[C.KEY]),
      comment_id: String(r[C.CID]),
      file_name: String(r[C.NAME]),
      participant: String(r[SCRIPT_COLS]),
      excerpt: String(r[C.EXCERPT]),
      code: code,
      theme: String(r[SCRIPT_COLS + 2]),
      memo: String(r[C.MEMO]),
      author: String(r[C.AUTHOR]),
      resolved: r[C.RESOLVED] === true,
      link: link ? link.replace(/""/g, '"') : '',
      colour: colours[code] || ''
    });
  });
  return json_({ rows: rows, generated: new Date().toISOString() });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad json' }); }
  if (!tokenOk_(body.token)) return json_({ error: 'bad token' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_({ error: 'a sync is running in the Sheet; try again in a minute' });
  try {
    ensureSheets_();
    return json_(applyBoard_(body.stickies || [], { by: body.by, boardId: body.boardId, boardName: body.boardName }));
  } catch (err) {
    return json_({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function tokenOk_(t) {
  var want = PropertiesService.getScriptProperties().getProperty('WEB_TOKEN');
  return !!want && t === want;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Canonical code → colour set in Codes!E (on the canonical code's own row). */
function codeColours_() {
  var sh = ss_().getSheetByName(SHEETS.CODES);
  var n = sh.getLastRow() - 1;
  var out = {};
  if (n < 1) return out;
  sh.getRange(2, 1, n, 5).getValues().forEach(function (r) {
    if (r[4]) out[String(r[0])] = String(r[4]).trim().toLowerCase();
  });
  return out;
}
