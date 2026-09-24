/**
 * Sheet layout and writes.
 *
 * Ownership rule: the script writes only Codings!A:M, Codes!A:B (new rows),
 * Sources!A:B and E:G, Log and Board. Everything else is the researcher's.
 */

var SHEETS = { CODINGS: 'Codings', CODES: 'Codes', SOURCES: 'Sources', THEMES: 'Themes',
  BOARD: 'Board history', RECODES: 'Recodes', BOARD_STATE: 'BoardState', LOG: 'Log' };

var CODINGS_HEADERS = ['row_key', 'comment_id', 'file_id', 'file_name', 'excerpt', 'code_raw',
  'memo', 'author', 'created', 'modified', 'resolved', 'deleted', 'source_type'];
var C = { KEY: 0, CID: 1, FILE: 2, NAME: 3, EXCERPT: 4, RAW: 5, MEMO: 6, AUTHOR: 7,
  CREATED: 8, MODIFIED: 9, RESOLVED: 10, DELETED: 11, TYPE: 12 };
var SCRIPT_COLS = CODINGS_HEADERS.length; // A–M

// Formula columns N–P. Row 1 holds header + ARRAYFORMULA so they fill themselves.
var CODINGS_FORMULAS = [
  '={"participant";ARRAYFORMULA(IF(A2:A="",,IFERROR(VLOOKUP(C2:C,Sources!A2:C,3,FALSE)&"","")))}',
  '={"code";ARRAYFORMULA(IF(A2:A="",,IF(IFERROR(VLOOKUP(F2:F,Codes!A2:B,2,FALSE)&"","")="",F2:F,VLOOKUP(F2:F,Codes!A2:B,2,FALSE)&"")))}',
  '={"theme";ARRAYFORMULA(IF(A2:A="",,IF(IFERROR(VLOOKUP(A2:A,BoardState!A2:B,2,FALSE)&"","")<>"",VLOOKUP(A2:A,BoardState!A2:B,2,FALSE)&"",IFERROR(VLOOKUP(O2:O,Codes!A2:C,3,FALSE)&"",IFERROR(VLOOKUP(F2:F,Codes!A2:C,3,FALSE)&"","")))))}'
];

// manual_theme is a fallback for codes not yet placed on the board; the board wins.
var CODES_HEADERS = ['code_raw', 'canonical', 'manual_theme', 'description', 'colour'];
var CODES_COUNT = '={"count";ARRAYFORMULA(IF(A2:A="",,COUNTIFS(Codings!F2:F,A2:A,Codings!L2:L,FALSE)))}';
var CODES_BOARD_COL = CODES_HEADERS.length + 2; // G, after count; written by Board.gs
var SOURCES_HEADERS = ['file_id', 'file_name', 'participant', 'include', 'last_synced', 'comment_rows', 'docx_parsed_at'];
var LOG_HEADERS = ['run_at', 'trigger', 'files_scanned', 'files_pending', 'rows_added', 'rows_updated', 'rows_marked_deleted', 'seconds', 'errors'];
var RECODES_HEADERS = ['time', 'by', 'file', 'excerpt', 'from', 'to', 'how'];
var BOARD_HEADERS = ['snapshot_time', 'synced_by', 'theme', 'code', 'participant', 'excerpt', 'row_key', 'x', 'y'];
var BOARD_STATE_HEADERS = ['row_key', 'theme', 'x', 'y', 'updated'];
var THEMES_HEADERS = ['theme', 'level', 'stickies', 'distinct_codes', 'codes', 'participants'];

function ss_() {
  // Bound script: the container is always the right book (menu, trigger or web app).
  return SpreadsheetApp.getActiveSpreadsheet() ||
    SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
}

/**
 * File → Make a copy duplicates this script *and* its saved settings. Anything
 * that points at the original (folder, web app, token, layout flags) is reset
 * so the copy starts clean instead of syncing into someone else's project.
 */
function resetIfCopied_(book) {
  var props = PropertiesService.getScriptProperties();
  var was = props.getProperty('SHEET_ID');
  if (was && was !== book.getId()) {
    ['FOLDER_ID', 'WEB_URL', 'WEB_TOKEN', 'LAYOUT_V1', 'LAYOUT_V2'].forEach(function (k) { props.deleteProperty(k); });
  }
  props.setProperty('SHEET_ID', book.getId());
}

/** Create any missing tabs, headers and formulas. Safe to run repeatedly. */
function ensureSheets_() {
  var book = ss_();
  resetIfCopied_(book);
  if (book.getSpreadsheetTimeZone() !== Session.getScriptTimeZone()) {
    book.setSpreadsheetTimeZone(Session.getScriptTimeZone());
  }
  function tab(name, headers) {
    var sh = book.getSheetByName(name) || book.insertSheet(name);
    if (sh.getRange(1, 1).getValue() === '') {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
    return sh;
  }
  function formulas(sh, col, fs) {
    // Rewritten every run: a formula that once pointed at a missing tab stays
    // #REF! even after the tab exists, and IFERROR would hide that as blanks.
    sh.getRange(1, col, 1, fs.length).setFormulas([fs]);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  }
  // The first Board tab had a vague "cluster" column; keep it, out of the way.
  var legacy = book.getSheetByName('Board');
  if (legacy && legacy.getRange(1, 3).getValue() === 'cluster') legacy.setName('Board (old)'); // deleted by tidyWorkbookOnce_

  // Lookup targets first, then the tabs whose formulas reference them.
  var sources = tab(SHEETS.SOURCES, SOURCES_HEADERS);
  var codes = tab(SHEETS.CODES, CODES_HEADERS);
  var state = tab(SHEETS.BOARD_STATE, BOARD_STATE_HEADERS);
  var codings = tab(SHEETS.CODINGS, CODINGS_HEADERS);
  tab(SHEETS.THEMES, THEMES_HEADERS);
  var history = tab(SHEETS.BOARD, BOARD_HEADERS);
  if (history.getRange(1, 2).getValue() === 'theme') {  // pre-"synced_by" layout
    history.insertColumnBefore(2);
    history.getRange(1, 2).setValue('synced_by');
  }
  tab(SHEETS.RECODES, RECODES_HEADERS);
  tab(SHEETS.LOG, LOG_HEADERS);
  formulas(codings, SCRIPT_COLS + 1, CODINGS_FORMULAS);
  formulas(codes, CODES_HEADERS.length + 1, [CODES_COUNT]);
  codes.getRange(1, 3).setValue('manual_theme');
  codes.getRange(1, CODES_BOARD_COL).setValue('board_themes').setFontWeight('bold');
  if (!state.isSheetHidden()) state.hideSheet(); // plumbing: the plugin's latest view
  codings.getRange('E:E').setWrap(true);
  codings.getRange('G:G').setWrap(true);
  codings.getRange('I:J').setNumberFormat('yyyy-mm-dd hh:mm');
  compactSources_(sources);
  applyLayoutOnce_(codings);
  tidyWorkbookOnce_(book);
  protectOnce_(book);
}

/**
 * Shared Sheet: warn (not block) anyone hand-editing what the sync writes,
 * since the next sync would overwrite it.
 */
function protectOnce_(book) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAYOUT_V3')) return;
  var codings = book.getSheetByName(SHEETS.CODINGS);
  codings.getRange(1, 1, codings.getMaxRows(), SCRIPT_COLS + 3).protect()
    .setDescription('Written by Sync. Change codes in the document or on the board; rename/merge in Codes.')
    .setWarningOnly(true);
  [SHEETS.THEMES, SHEETS.BOARD, SHEETS.RECODES, SHEETS.BOARD_STATE, SHEETS.LOG].forEach(function (name) {
    book.getSheetByName(name).protect().setDescription('Written by Sync.').setWarningOnly(true);
  });
  var recodes = book.getSheetByName(SHEETS.RECODES);
  if (!recodes.isSheetHidden()) recodes.hideSheet();
  props.setProperty('LAYOUT_V3', '1');
}

/**
 * One-time view tweaks. Guarded by a property so a column the researcher
 * unhides stays visible.
 */
function applyLayoutOnce_(codings) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAYOUT_V1')) return;
  // Plumbing columns: row_key, comment_id, file_id, created … source_type.
  [[C.KEY, 3], [C.CREATED, 5]].forEach(function (span) {
    codings.hideColumns(span[0] + 1, span[1]);
  });
  // With `deleted` hidden, grey out and strike through deleted rows instead.
  var rules = codings.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$L2=TRUE')
    .setFontColor('#9e9e9e')
    .setStrikethrough(true)
    .setRanges([codings.getRange('A2:P')])
    .build());
  codings.setConditionalFormatRules(rules);
  props.setProperty('LAYOUT_V1', '1');
}

/**
 * One-time tidy of the workbook: drop the blank starter tab and the pre-themes
 * Board tab, put the working tabs first, hide the record-keeping ones.
 */
function tidyWorkbookOnce_(book) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAYOUT_V2')) return;
  var ours = {};
  Object.keys(SHEETS).forEach(function (k) { ours[SHEETS[k]] = true; });
  book.getSheets().forEach(function (sh) {
    var blankStarter = !ours[sh.getName()] && sh.getLastRow() === 0 && sh.getLastColumn() === 0;
    if (blankStarter || sh.getName() === 'Board (old)') book.deleteSheet(sh);
  });
  [SHEETS.CODINGS, SHEETS.THEMES, SHEETS.CODES, SHEETS.SOURCES, SHEETS.BOARD, SHEETS.RECODES, SHEETS.LOG, SHEETS.BOARD_STATE]
    .forEach(function (name, i) {
      var sh = book.getSheetByName(name);
      if (!sh) return;
      book.setActiveSheet(sh);
      book.moveActiveSheet(i + 1);
    });
  [SHEETS.BOARD, SHEETS.LOG, SHEETS.BOARD_STATE].forEach(function (name) {
    var sh = book.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });
  book.setActiveSheet(book.getSheetByName(SHEETS.CODINGS));
  props.setProperty('LAYOUT_V2', '1');
}

/**
 * Force strings to stay text: otherwise a transcript line starting with = + - @
 * becomes a formula, and "2020" or "3/4" become a number or date.
 */
function asText_(v) {
  return typeof v === 'string' && v !== '' ? "'" + v : v;
}

function readSources_() {
  var sh = ss_().getSheetByName(SHEETS.SOURCES);
  var n = sh.getLastRow() - 1;
  var map = {};
  if (n < 1) return map;
  sh.getRange(2, 1, n, SOURCES_HEADERS.length).getValues().forEach(function (r, i) {
    if (!r[0]) return;
    map[r[0]] = {
      row: i + 2, name: r[1], participant: r[2], include: r[3] !== false,
      lastSynced: r[4] instanceof Date ? r[4] : null,
      docxParsedAt: r[6] instanceof Date ? r[6] : null
    };
  });
  return map;
}

/**
 * Drop rows with no file_id (e.g. stray checkboxes) and close the gaps.
 * Checkboxes count as content, so they must only exist on real rows.
 */
function compactSources_(sh) {
  var n = sh.getMaxRows() - 1;
  if (n < 1) return;
  var range = sh.getRange(2, 1, n, SOURCES_HEADERS.length);
  var values = range.getValues();
  var junk = values.some(function (r) {
    return r[0] === '' && r.some(function (v) { return v !== ''; });
  });
  if (!junk) return;
  var keep = values.filter(function (r) { return r[0] !== ''; });
  range.clearDataValidations().clearContent();
  if (keep.length) {
    sh.getRange(2, 1, keep.length, SOURCES_HEADERS.length)
      .setValues(keep.map(function (r) { return r.map(asText_); }));
    sh.getRange(2, 4, keep.length, 1).insertCheckboxes();
  }
}

/** Write back sync state; new files are appended with include ticked. */
function writeSources_(files, sources, counts) {
  var sh = ss_().getSheetByName(SHEETS.SOURCES);
  files.forEach(function (f) {
    var s = sources[f.id];
    if (!s || !s.touched) return;
    var state = [[s.lastSynced || '', counts[f.id] || 0, s.docxParsedAt || '']];
    if (s.row) {
      sh.getRange(s.row, 2).setValue(asText_(f.name));
      sh.getRange(s.row, 5, 1, 3).setValues(state);
    } else {
      s.row = sh.getLastRow() + 1;
      sh.getRange(s.row, 1, 1, SOURCES_HEADERS.length)
        .setValues([[asText_(f.id), asText_(f.name), '', true].concat(state[0])]);
      sh.getRange(s.row, 4).insertCheckboxes();
    }
  });
}

/**
 * Apply one run's comments to Codings.
 * @param {Array<Object>} comments normalised comments from any source
 * @param {Array<string>} fullyScannedDocx file IDs whose embedded comments were
 *        read in full: any embedded row for those files not seen is now gone.
 */
function upsertCodings_(comments, fullyScannedDocx) {
  var sh = ss_().getSheetByName(SHEETS.CODINGS);
  var n = sh.getLastRow() - 1;
  var data = n > 0 ? sh.getRange(2, 1, n, SCRIPT_COLS).getValues() : [];
  if (n > 0) {
    // Keep D's HYPERLINK formulas instead of their display text.
    sh.getRange(2, C.NAME + 1, n, 1).getFormulas().forEach(function (f, i) {
      if (f[0]) data[i][C.NAME] = f[0];
    });
  }
  var byKey = {};
  var byComment = {};
  data.forEach(function (r, i) {
    byKey[r[C.KEY]] = i;
    (byComment[r[C.CID]] = byComment[r[C.CID]] || []).push(i);
  });

  var stats = { added: 0, updated: 0, deleted: 0 };
  var changed = {};
  var seenKeys = {};
  var newCodes = {};

  function markDeleted(i) {
    if (data[i][C.DELETED] !== true) { data[i][C.DELETED] = true; changed[i] = true; stats.deleted++; }
  }

  comments.forEach(function (c) {
    var liveKeys = {};
    if (!c.deleted) {
      commentToRows(c).forEach(function (row) {
        liveKeys[row[C.KEY]] = true;
        seenKeys[row[C.KEY]] = true;
        if (row[C.RAW]) newCodes[row[C.RAW]] = true;
        var i = byKey[row[C.KEY]];
        if (i === undefined) {
          byKey[row[C.KEY]] = data.length;
          data.push(row);
          stats.added++;
        } else if (!sameRow_(data[i], row)) {
          data[i] = row;
          changed[i] = true;
          stats.updated++;
        }
      });
    }
    // Codes removed from the comment (or the whole comment deleted): keep the row, flag it.
    (byComment[c.key] || []).forEach(function (i) {
      if (!liveKeys[data[i][C.KEY]]) markDeleted(i);
    });
  });

  var gone = {};
  (fullyScannedDocx || []).forEach(function (id) { gone[id] = true; });
  data.forEach(function (r, i) {
    if (gone[r[C.FILE]] && r[C.TYPE] === 'docx-embedded' && !seenKeys[r[C.KEY]]) markDeleted(i);
  });

  var out = data.map(function (r) {
    return r.map(function (v, j) { return j === C.NAME ? v : asText_(v); });
  });
  if (Object.keys(changed).length && n > 0) {
    sh.getRange(2, 1, n, SCRIPT_COLS).setValues(out.slice(0, n));
  }
  if (out.length > n) {
    sh.getRange(n + 2, 1, out.length - n, SCRIPT_COLS).setValues(out.slice(n));
  }
  appendNewCodes_(Object.keys(newCodes));
  return stats;
}

function sameRow_(a, b) {
  for (var j = 0; j < SCRIPT_COLS; j++) {
    var x = a[j] instanceof Date ? a[j].getTime() : a[j];
    var y = b[j] instanceof Date ? b[j].getTime() : b[j];
    if (String(x) !== String(y)) return false;
  }
  return true;
}

function appendNewCodes_(codes) {
  if (!codes.length) return;
  var sh = ss_().getSheetByName(SHEETS.CODES);
  var n = sh.getLastRow() - 1;
  var have = {};
  if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function (r) { have[String(r[0])] = true; });
  var add = codes.filter(function (c) { return !have[c]; }).sort()
    .map(function (c) { return [asText_(c), asText_(c)]; });
  if (add.length) sh.getRange(n + 2, 1, add.length, 2).setValues(add);
}

/** Codings row count per file (live rows only), for the Sources sheet. */
function countRowsByFile_() {
  var sh = ss_().getSheetByName(SHEETS.CODINGS);
  var n = sh.getLastRow() - 1;
  var counts = {};
  if (n < 1) return counts;
  sh.getRange(2, 1, n, SCRIPT_COLS).getValues().forEach(function (r) {
    if (r[C.DELETED] !== true) counts[r[C.FILE]] = (counts[r[C.FILE]] || 0) + 1;
  });
  return counts;
}

function log_(entry) {
  ss_().getSheetByName(SHEETS.LOG).appendRow([
    new Date(), entry.trigger, entry.scanned, entry.pending, entry.added, entry.updated,
    entry.deleted, entry.seconds, entry.errors.join(' | ').slice(0, 45000)
  ]);
}
