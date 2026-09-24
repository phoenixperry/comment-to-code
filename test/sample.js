// Parse sample-data/P99_sample_interview.docx end-to-end (unzip → extract) and print rows.
const path = require('path');
const { execFileSync } = require('child_process');
const crypto = require('crypto');
const { DOMParser } = require('@xmldom/xmldom');
const gs = (f) => require(path.join(__dirname, '..', 'apps-script', f));
const { extractDocxComments } = gs('Docx.gs');
const { commentToRows } = gs('Parse.gs');

const file = path.join(__dirname, '..', 'sample-data', 'P99_sample_interview.docx');
const read = (p) => { try { return execFileSync('unzip', ['-p', file, p]).toString('utf8'); } catch { return undefined; } };
const parts = {};
for (const p of ['word/document.xml', 'word/comments.xml', 'word/commentsExtended.xml']) parts[p] = read(p);

const A = {
  parse: (s) => new DOMParser().parseFromString(s, 'text/xml').documentElement,
  name: (el) => el.localName,
  attr: (el, l) => { for (const a of Array.from(el.attributes)) if (a.localName === l) return a.value; return null; },
  kids: (el) => Array.from(el.childNodes).filter((n) => n.nodeType === 1),
  text: (el) => el.textContent,
};
const sha = (s) => crypto.createHash('sha1').update(s).digest('hex').slice(0, 12);
const comments = extractDocxComments(parts, A, sha).map((c) => ({
  ...c, key: 'FILE:' + c.localKey, fileId: 'FILE', fileName: 'P99_sample_interview.docx',
  link: 'https://drive.google.com/file/d/FILE', modified: c.created, deleted: false, sourceType: 'docx-embedded',
}));
for (const c of comments) for (const r of commentToRows(c))
  console.log([r[5] || '(uncoded)', JSON.stringify(r[4]), JSON.stringify(r[6]), 'resolved=' + r[10]].join('  |  '));
