/**
 * Comments embedded in .docx files (made in Word desktop, then uploaded).
 *
 * The parsing core (extractDocxComments) works through a tiny XML adapter,
 * so it runs on XmlService here and on @xmldom/xmldom in Node tests.
 */

var XmlServiceAdapter = {
  parse: function (s) { return XmlService.parse(s).getRootElement(); },
  name: function (el) { return el.getName(); },
  attr: function (el, local) {
    var attrs = el.getAttributes();
    for (var i = 0; i < attrs.length; i++) {
      if (attrs[i].getName() === local) return attrs[i].getValue();
    }
    return null;
  },
  kids: function (el) { return el.getChildren(); },
  text: function (el) { return el.getText(); }
};

/** Apps Script entry point: one Drive file → normalised comments. */
function fetchDocxEmbeddedComments(file) {
  var blob = DriveApp.getFileById(file.id).getBlob().setContentType('application/zip');
  var parts = {};
  Utilities.unzip(blob).forEach(function (b) {
    var n = b.getName();
    if (n === 'word/document.xml' || n === 'word/comments.xml' || n === 'word/commentsExtended.xml') {
      parts[n] = b.getDataAsString('UTF-8');
    }
  });
  if (!parts['word/comments.xml']) return [];
  var found = extractDocxComments(parts, XmlServiceAdapter, sha1Short_);
  return found.map(function (c) {
    return {
      key: file.id + ':' + c.localKey,
      fileId: file.id,
      fileName: file.name,
      link: file.url,
      excerpt: c.excerpt || '[no highlighted range]',
      content: c.content,
      author: c.author,
      created: c.created,
      modified: c.created,
      resolved: c.resolved,
      deleted: false,
      replies: c.replies,
      sourceType: 'docx-embedded'
    };
  });
}

function sha1Short_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); })
    .join('').slice(0, 12);
}

/**
 * @param {Object<string,string>} parts XML strings keyed by zip path
 * @param {Object} A XML adapter
 * @param {function(string):string} hash
 * @return {Array<{localKey,author,created,content,excerpt,resolved,replies}>}
 */
function extractDocxComments(parts, A, hash) {
  var excerpts = parts['word/document.xml']
    ? collectRanges_(A.parse(parts['word/document.xml']), A) : {};

  // commentsExtended links replies to parents (via paragraph IDs) and holds "done" state.
  var ext = {};
  if (parts['word/commentsExtended.xml']) {
    A.kids(A.parse(parts['word/commentsExtended.xml'])).forEach(function (e) {
      if (A.name(e) !== 'commentEx') return;
      ext[A.attr(e, 'paraId')] = { parent: A.attr(e, 'paraIdParent'), done: A.attr(e, 'done') === '1' };
    });
  }

  var all = A.kids(A.parse(parts['word/comments.xml']))
    .filter(function (el) { return A.name(el) === 'comment'; })
    .map(function (el) {
      var paras = descendants_(el, A, 'p');
      var last = paras[paras.length - 1];
      return {
        id: A.attr(el, 'id'),
        author: A.attr(el, 'author') || '',
        created: A.attr(el, 'date') || '',
        content: paras.map(function (p) {
          return descendants_(p, A, 't').map(A.text).join('');
        }).join('\n').trim(),
        paraId: last ? A.attr(last, 'paraId') : null
      };
    });

  var byPara = {};
  all.forEach(function (c) { if (c.paraId) byPara[c.paraId] = c; });

  var top = [];
  var used = {};
  all.forEach(function (c) {
    var info = c.paraId ? ext[c.paraId] : null;
    var parent = info && info.parent ? byPara[info.parent] : null;
    if (parent) {
      (parent.replies = parent.replies || []).push({
        content: c.content, author: c.author, created: c.created, deleted: false
      });
      return;
    }
    // Key on author + creation time: stable when the comment text is edited or
    // the file is re-saved (w:id renumbers on save). Fall back to content when
    // Word has stripped dates ("remove personal information").
    var excerpt = (excerpts[c.id] || '').trim();
    var base = hash(c.created ? c.author + '|' + c.created : c.author + '|' + c.content + '|' + excerpt);
    var key = base;
    for (var n = 2; used[key]; n++) key = base + '~' + n;
    used[key] = true;
    top.push({
      localKey: key,
      wId: c.id,
      author: c.author,
      created: c.created,
      content: c.content,
      excerpt: excerpt,
      resolved: !!(info && info.done),
      replies: c.replies = c.replies || []
    });
  });
  return top;
}

/**
 * Swap one #tag inside comment w:id=wId, editing only the text of a single <w:t>.
 * Pure string surgery so every other byte of the XML is untouched.
 * @return {?string} new comments.xml, or null if the tag wasn't found in one run
 */
function rewriteDocxCommentTag(commentsXml, wId, oldTag, newTag) {
  var open = new RegExp('<w:comment\\b[^>]*\\bw:id="' + wId + '"[^>]*>');
  var m = open.exec(commentsXml);
  if (!m) return null;
  var start = m.index + m[0].length;
  var end = commentsXml.indexOf('</w:comment>', start);
  if (end < 0) return null;
  var body = commentsXml.slice(start, end);
  var done = false;
  var newBody = body.replace(/<w:t(\s[^>]*)?>([^<]*)<\/w:t>/g, function (all, attrs, text) {
    if (done) return all;
    var next = replaceTagInText(text, oldTag, newTag);
    if (next === null) return all;
    done = true;
    return '<w:t xml:space="preserve">' + next + '</w:t>';
  });
  if (!done && !oldTag) {
    // Uncoded comment with no text runs: nothing sensible to edit.
    return null;
  }
  return done ? commentsXml.slice(0, start) + newBody + commentsXml.slice(end) : null;
}

/** Apps Script: recode an embedded Word comment and upload the file as a new version. */
function writeDocxRecode_(fileId, localKey, oldTag, newTag) {
  var file = DriveApp.getFileById(fileId);
  var blobs = Utilities.unzip(file.getBlob().setContentType('application/zip'));
  var parts = {};
  blobs.forEach(function (b) { parts[b.getName()] = b.getDataAsString('UTF-8'); });
  var found = extractDocxComments(parts, XmlServiceAdapter, sha1Short_)
    .filter(function (c) { return c.localKey === localKey; })[0];
  if (!found) throw new Error('comment not found in file');
  var xml = rewriteDocxCommentTag(parts['word/comments.xml'], found.wId, oldTag, newTag);
  if (xml === null) throw new Error('#' + (oldTag || '(none)') + ' not found as plain text in the Word comment');
  var out = blobs.map(function (b) {
    return b.getName() === 'word/comments.xml'
      ? Utilities.newBlob(xml, 'application/xml', 'word/comments.xml') : b;
  });
  var zipped = Utilities.zip(out, file.getName())
    .setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  Drive.Files.update({}, fileId, zipped);  // Drive keeps the previous version
}

/** Walk document.xml in order, appending text to every comment range that is open. */
function collectRanges_(root, A) {
  var open = {};
  var out = {};
  function add(s) { for (var id in open) out[id] += s; }
  (function walk(el) {
    switch (A.name(el)) {
      case 'commentRangeStart':
        var sid = A.attr(el, 'id'); open[sid] = true; if (!(sid in out)) out[sid] = ''; return;
      case 'commentRangeEnd':
        delete open[A.attr(el, 'id')]; return;
      case 'delText': case 'instrText': return;  // tracked deletions, field codes
      case 't': add(A.text(el)); return;
      case 'tab': add('\t'); return;
      case 'br': case 'cr': add('\n'); return;
    }
    A.kids(el).forEach(walk);
    if (A.name(el) === 'p') add('\n');
  })(root);
  for (var id in out) out[id] = out[id].replace(/\n+$/, '').replace(/\n{2,}/g, '\n');
  return out;
}

function descendants_(el, A, local) {
  var out = [];
  (function walk(e) {
    A.kids(e).forEach(function (k) { if (A.name(k) === local) out.push(k); walk(k); });
  })(el);
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = { extractDocxComments: extractDocxComments, rewriteDocxCommentTag: rewriteDocxCommentTag };
}
