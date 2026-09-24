/**
 * Board → Sheet → documents.
 *
 * The FigJam plugin posts every coded sticky: its row_key, the theme path it
 * sits in ("Studio as refuge › Failure as ritual", or "" if unthemed), the tag
 * currently written on it, and its position. From that we:
 *   1. write recoded tags back into the source comments (then re-read those files)
 *   2. record the board's current state (BoardState) and a dated history (Board)
 *   3. rebuild the Themes summary and Codes!board_themes
 */

var THEME_SEP = ' › ';

function applyBoard_(stickies, meta) {
  meta = meta || {};
  // One board per Sheet: a second board would overwrite the first one's themes.
  var props = PropertiesService.getScriptProperties();
  var linked = props.getProperty('BOARD_ID');
  if (meta.boardId && linked && linked !== meta.boardId) {
    return { error: 'This Sheet is linked to another board ("' + props.getProperty('BOARD_NAME') +
      '"). Sync from that board, or use Coding → Unlink FigJam board in the Sheet to switch.', fatal: true };
  }
  if (meta.boardId && !linked) {
    props.setProperty('BOARD_ID', meta.boardId);
    props.setProperty('BOARD_NAME', meta.boardName || meta.boardId);
  }
  // An empty page (e.g. the plugin run on a fresh board) must not wipe saved themes.
  if (!stickies.length) return { ok: true, written: 0, rekeys: {}, errors: [] };
  var by = String(meta.by || 'someone');
  var recodeLog = [];
  var errors = [];
  var rekeys = {};
  var rows = codingsIndex_();

  // 1. Recodes: tag on the sticky differs from the row's current code.
  var touched = {};
  stickies.forEach(function (st) {
    var tag = normaliseTag(st.tag);
    var row = rows[st.row_key];
    if (!tag || !row || row.deleted || tag === row.code) return;
    try {
      var how = writeRecode_(row, tag, by);
      rekeys[st.row_key] = row.commentId + '#' + tag;
      touched[row.fileId] = true;
      recodeLog.push([new Date(), asText_(by), asText_(row.fileName), asText_(clip_(row.excerpt, 120)),
        asText_('#' + (row.code || '(uncoded)')), asText_('#' + tag), how]);
    } catch (err) {
      errors.push((row.fileName || row.fileId) + ' (#' + row.code + ' → #' + tag + '): ' + err.message);
    }
  });
  if (Object.keys(touched).length) {
    syncFiles_(Object.keys(touched));
    SpreadsheetApp.flush();
    rows = codingsIndex_();
  }

  if (recodeLog.length) {
    var rl = ss_().getSheetByName(SHEETS.RECODES);
    rl.getRange(rl.getLastRow() + 1, 1, recodeLog.length, RECODES_HEADERS.length).setValues(recodeLog);
  }

  // 2. Current state + history, keyed by the post-recode row_key.
  var now = new Date();
  var state = stickies.map(function (st) {
    return { key: rekeys[st.row_key] || st.row_key, theme: cleanTheme_(st.theme), x: st.x, y: st.y };
  });
  writeBoardState_(state, now);
  appendBoardHistory_(state, rows, now, by);

  // 3. Summaries.
  writeThemes_(state, rows);
  writeCodeThemes_(state, rows);
  return { ok: true, written: state.length, rekeys: rekeys, errors: errors };
}

/**
 * Recode one row in its source document. Runs as whoever deployed the web app,
 * so `by` (the researcher at the board) is recorded wherever it can be.
 * @return {string} how it was written, for the Recodes log
 */
function writeRecode_(row, newTag, by) {
  var oldTag = row.codeRaw;
  if (row.sourceType === 'docx-embedded') {
    var localKey = row.commentId.slice(row.commentId.indexOf(':') + 1);
    writeDocxRecode_(row.fileId, localKey, oldTag, newTag);
    return 'Word comment edited (new file version)';
  }
  var c = Drive.Comments.get(row.fileId, row.commentId, { fields: 'content,author(me)' });
  var edited = c.author && c.author.me ? replaceTagInText(c.content, oldTag, newTag) : null;
  if (edited !== null) {
    Drive.Comments.update({ content: edited }, row.fileId, row.commentId, { fields: 'id' });
    return 'comment edited';
  }
  // Someone else's comment (or a tag that only exists via an earlier recode reply):
  // leave their words alone and record the recode as a reply the parser understands.
  Drive.Replies.create({ content: '↻ #' + oldTag + ' → #' + newTag + ' (' + by + ', via board)' },
    row.fileId, row.commentId, { fields: 'id' });
  return 'reply added';
}

function clip_(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

function cleanTheme_(t) {
  return String(t || '').split(THEME_SEP).map(function (s) { return s.trim(); }).filter(String).join(THEME_SEP);
}

/** row_key → the fields the board logic needs (reads formula columns N–O too). */
function codingsIndex_() {
  var sh = ss_().getSheetByName(SHEETS.CODINGS);
  var n = sh.getLastRow() - 1;
  var out = {};
  if (n < 1) return out;
  sh.getRange(2, 1, n, SCRIPT_COLS + 2).getValues().forEach(function (r) {
    if (!r[C.KEY]) return;
    out[r[C.KEY]] = {
      commentId: String(r[C.CID]),
      fileId: String(r[C.FILE]),
      fileName: String(r[C.NAME]),
      excerpt: String(r[C.EXCERPT]),
      codeRaw: String(r[C.RAW]),
      sourceType: String(r[C.TYPE]),
      deleted: r[C.DELETED] === true,
      participant: String(r[SCRIPT_COLS]),
      code: String(r[SCRIPT_COLS + 1])
    };
  });
  return out;
}

function writeBoardState_(state, now) {
  var sh = ss_().getSheetByName(SHEETS.BOARD_STATE);
  var n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, BOARD_STATE_HEADERS.length).clearContent();
  if (!state.length) return;
  sh.getRange(2, 1, state.length, BOARD_STATE_HEADERS.length).setValues(state.map(function (s) {
    return [asText_(s.key), asText_(s.theme), Math.round(s.x) || 0, Math.round(s.y) || 0, now];
  }));
}

function appendBoardHistory_(state, rows, now, by) {
  if (!state.length) return;
  var sh = ss_().getSheetByName(SHEETS.BOARD);
  sh.getRange(sh.getLastRow() + 1, 1, state.length, BOARD_HEADERS.length).setValues(state.map(function (s) {
    var r = rows[s.key] || {};
    var excerpt = r.excerpt || '';
    return [now, asText_(by), asText_(s.theme), asText_(r.code || ''), asText_(r.participant || ''),
      asText_(clip_(excerpt, 80)), asText_(s.key), Math.round(s.x) || 0, Math.round(s.y) || 0];
  }));
}

/** One row per theme and per sub-theme; a sticky in "A › B" counts toward A and A › B. */
function writeThemes_(state, rows) {
  var agg = {};
  state.forEach(function (s) {
    if (!s.theme) return;
    var r = rows[s.key];
    if (!r || r.deleted) return;
    var parts = s.theme.split(THEME_SEP);
    for (var d = 1; d <= parts.length; d++) {
      var path = parts.slice(0, d).join(THEME_SEP);
      var a = agg[path] || (agg[path] = { level: d, stickies: 0, codes: {}, people: {} });
      a.stickies++;
      var code = r.code || '(uncoded)';
      a.codes[code] = (a.codes[code] || 0) + 1;
      if (r.participant) a.people[r.participant] = true;
    }
  });
  var out = Object.keys(agg).sort().map(function (path) {
    var a = agg[path];
    return [asText_(path), a.level, a.stickies, Object.keys(a.codes).length,
      asText_(tally_(a.codes)), asText_(Object.keys(a.people).sort().join(', '))];
  });
  var sh = ss_().getSheetByName(SHEETS.THEMES);
  var n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, THEMES_HEADERS.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, THEMES_HEADERS.length).setValues(out);
}

/** Codes!board_themes: where each code's stickies sit, e.g. "Studio as refuge ×3 · Money talk ×1". */
function writeCodeThemes_(state, rows) {
  var byCode = {};
  state.forEach(function (s) {
    var r = rows[s.key];
    if (!r || r.deleted || !r.code) return;
    var t = s.theme || '(unthemed)';
    var m = byCode[r.code] || (byCode[r.code] = {});
    m[t] = (m[t] || 0) + 1;
  });
  var sh = ss_().getSheetByName(SHEETS.CODES);
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var canon = sh.getRange(2, 1, n, 2).getValues();
  sh.getRange(2, CODES_BOARD_COL, n, 1).setValues(canon.map(function (r) {
    var code = String(r[1] || r[0]);
    return [byCode[code] ? asText_(tally_(byCode[code], ' · ')) : ''];
  }));
}

/** {a:3,b:1} → "a ×3, b ×1", most frequent first. */
function tally_(counts, sep) {
  return Object.keys(counts)
    .sort(function (x, y) { return counts[y] - counts[x] || x.localeCompare(y); })
    .map(function (k) { return k + ' ×' + counts[k]; })
    .join(sep || ', ');
}
