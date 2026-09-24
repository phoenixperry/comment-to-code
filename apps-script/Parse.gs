/**
 * Comment text → codes + memo. Pure functions, no Google services,
 * so they also run under Node for tests (see test/run.js).
 */

// A #tag must follow start-of-text or a non-word, non-URL character,
// so "example.com/page#section" and "&#39;" are not codes. Codes start with a letter
// (so "#1" in a numbered list is ignored). "/" makes a subcode: #failure/ritual.
var TAG_RE = /(^|[^\p{L}\p{N}_\/&#])#(\p{L}[\p{L}\p{N}_\/-]*)/gu;

// A reply that recodes the comment it answers, written by the board sync when
// the comment belongs to someone else and can't be edited: "↻ #mood → #atmosphere".
// An empty old tag ("↻ # → #x") adds a code to an uncoded comment.
var RECODE_RE = /^\s*↻\s*#([\p{L}\p{N}_\/-]*)\s*→\s*#(\p{L}[\p{L}\p{N}_\/-]*)/u;

function normaliseTag(tag) {
  return String(tag || '').replace(/[-_\/]+$/, '').replace(/\/{2,}/g, '/').toLowerCase();
}

function parseComment(text) {
  text = String(text || '');
  var codes = [];
  var seen = {};
  text.replace(TAG_RE, function (_, pre, tag) {
    var code = normaliseTag(tag);
    if (code && !seen[code]) { seen[code] = true; codes.push(code); }
    return '';
  });
  var memo = text
    .replace(TAG_RE, '$1')
    .replace(/[ \t]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return { codes: codes.length ? codes : [''], memo: memo };
}

/** Apply "↻ #old → #new" replies, oldest first, to a code list. */
function applyRecodes(codes, replies) {
  var out = codes.filter(String);
  (replies || []).forEach(function (r) {
    var m = !r.deleted && RECODE_RE.exec(r.content || '');
    if (!m) return;
    var from = normaliseTag(m[1]);
    var to = normaliseTag(m[2]);
    var i = out.indexOf(from);
    if (from && i >= 0) out.splice(i, 1);
    if (out.indexOf(to) < 0) out.splice(i >= 0 ? i : out.length, 0, to);
  });
  return out.length ? out : [''];
}

function isRecodeReply(r) { return RECODE_RE.test(r.content || ''); }

/** Replies → memo lines, oldest first. Deleted and recode replies are skipped. */
function formatReplies(replies) {
  return (replies || [])
    .filter(function (r) { return !r.deleted && r.content && !isRecodeReply(r); })
    .map(function (r) {
      var day = r.created ? String(r.created).slice(0, 10) : '';
      return '↳ ' + (r.author || '?') + (day ? ' (' + day + ')' : '') + ': ' + r.content.trim();
    })
    .join('\n');
}

/**
 * One normalised comment → Codings rows (one per code).
 * Column order must match CODINGS_HEADERS in Sheet.gs (script-owned A–M).
 */
function commentToRows(c) {
  var parsed = parseComment(c.content);
  var codes = applyRecodes(parsed.codes, c.replies);
  var replyText = formatReplies(c.replies);
  var memo = [parsed.memo, replyText].filter(String).join('\n');
  return codes.map(function (code) {
    return [
      c.key + '#' + code,     // A row_key
      c.key,                  // B comment_id
      c.fileId,               // C file_id
      c.link ? hyperlink_(c.link, c.fileName) : c.fileName, // D file_name
      c.excerpt || '',        // E excerpt
      code,                   // F code_raw
      memo,                   // G memo
      c.author || '',         // H author
      c.created ? new Date(c.created) : '', // I created
      c.modified ? new Date(c.modified) : '', // J modified
      !!c.resolved,           // K resolved
      !!c.deleted,            // L deleted
      c.sourceType            // M source_type
    ];
  });
}

/**
 * Replace one #tag in comment text, keeping everything else as written.
 * Empty oldTag adds the new tag at the front. Returns null if oldTag isn't there.
 */
function replaceTagInText(text, oldTag, newTag) {
  text = String(text || '');
  if (!oldTag) return '#' + newTag + (text ? ' ' + text : '');
  var re = new RegExp('(^|[^\\p{L}\\p{N}_\\/&#])#' + escapeRe_(oldTag) + '(?![\\p{L}\\p{N}_\\/-])', 'iu');
  if (!re.test(text)) return null;
  return text.replace(re, function (_, pre) { return pre + '#' + newTag; });
}

function escapeRe_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** Drive returns quoted excerpts HTML-escaped ("you&#39;re"). */
function decodeEntities(s) {
  var named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0' };
  return String(s || '').replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, function (m, e) {
    if (e.charAt(0) === '#') {
      var code = e.charAt(1).toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return isNaN(code) ? m : String.fromCodePoint(code);
    }
    return named.hasOwnProperty(e.toLowerCase()) ? named[e.toLowerCase()] : m;
  });
}

function hyperlink_(url, label) {
  var q = function (s) { return String(s).replace(/"/g, '""'); };
  return '=HYPERLINK("' + q(url) + '","' + q(label) + '")';
}

if (typeof module !== 'undefined') {
  module.exports = { replaceTagInText: replaceTagInText, applyRecodes: applyRecodes, normaliseTag: normaliseTag, decodeEntities: decodeEntities, parseComment: parseComment, formatReplies: formatReplies, commentToRows: commentToRows };
}
