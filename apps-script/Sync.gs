/**
 * One sync run: walk the watched folder, pull new/changed comments, write the sheet.
 * Runs from the menu or the 10-minute trigger.
 */

var TIME_BUDGET_MS = 4.5 * 60 * 1000;  // Apps Script kills runs at 6 min
var SINCE_OVERLAP_MS = 2 * 60 * 1000;  // re-read a little overlap; upserts are idempotent
var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function syncAll(e) {
  var trigger = e && e.triggerUid ? 'auto' : 'manual';
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { skipped: true };
  var started = Date.now();
  var errors = [];
  try {
    ensureSheets_();
    var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    if (!folderId) throw new Error('No watched folder. Use Coding → Set watched folder…');

    var files = listFiles_(DriveApp.getFolderById(folderId));
    var sources = readSources_();
    // Least recently synced first, so a run that hits the time limit never starves a file.
    files.sort(function (a, b) {
      var ta = sources[a.id] && sources[a.id].lastSynced ? sources[a.id].lastSynced.getTime() : 0;
      var tb = sources[b.id] && sources[b.id].lastSynced ? sources[b.id].lastSynced.getTime() : 0;
      return ta - tb;
    });

    var comments = [];
    var fullDocx = [];
    var scanned = 0;
    for (var i = 0; i < files.length; i++) {
      if (Date.now() - started > TIME_BUDGET_MS) break;
      var f = files[i];
      var s = sources[f.id] || (sources[f.id] = { include: true });
      s.touched = true;
      if (!s.include) continue;
      try {
        comments = comments.concat(fileComments_(f, s, fullDocx));
      } catch (err) {
        errors.push(f.name + ': ' + err.message);
      }
      scanned++;
    }

    var stats = upsertCodings_(comments, fullDocx);
    writeSources_(files, sources, countRowsByFile_());
    log_({
      trigger: trigger, scanned: scanned, pending: files.length - i,
      added: stats.added, updated: stats.updated, deleted: stats.deleted,
      seconds: Math.round((Date.now() - started) / 1000), errors: errors
    });
    return stats;
  } catch (err) {
    errors.push(err.message);
    try { log_({ trigger: trigger, scanned: 0, pending: '', added: 0, updated: 0, deleted: 0,
      seconds: Math.round((Date.now() - started) / 1000), errors: errors }); } catch (_) {}
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** New/changed comments for one file; updates its Sources state `s` in memory. */
function fileComments_(f, s, fullDocx) {
  var runAt = new Date();
  var since = s.lastSynced ? new Date(s.lastSynced.getTime() - SINCE_OVERLAP_MS) : null;
  var out = fetchDriveComments(f, since);
  if (f.sourceType === 'docx-drive' && (!s.docxParsedAt || f.updated > s.docxParsedAt)) {
    out = out.concat(fetchDocxEmbeddedComments(f));
    fullDocx.push(f.id);
    s.docxParsedAt = runAt;
  }
  s.lastSynced = runAt;
  return out;
}

/** Re-read specific files now (after a recode). Caller holds the script lock. */
function syncFiles_(fileIds) {
  var sources = readSources_();
  var files = fileIds.map(function (id) { return fileInfo_(DriveApp.getFileById(id)); });
  var comments = [];
  var fullDocx = [];
  files.forEach(function (f) {
    var s = sources[f.id] || (sources[f.id] = { include: true });
    s.touched = true;
    comments = comments.concat(fileComments_(f, s, fullDocx));
  });
  upsertCodings_(comments, fullDocx);
  writeSources_(files, sources, countRowsByFile_());
}

function fileInfo_(f) {
  return {
    id: f.getId(),
    name: f.getName(),
    url: f.getUrl(),
    updated: f.getLastUpdated(),
    sourceType: f.getMimeType() === DOCX_MIME ? 'docx-drive' : 'gdoc'
  };
}

/** Google Docs and .docx files in the folder and all subfolders. */
function listFiles_(folder) {
  var out = [];
  (function walk(fo) {
    [MimeType.GOOGLE_DOCS, DOCX_MIME].forEach(function (mime) {
      var it = fo.getFilesByType(mime);
      while (it.hasNext()) out.push(fileInfo_(it.next()));
    });
    var subs = fo.getFolders();
    while (subs.hasNext()) walk(subs.next());
  })(folder);
  return out;
}
