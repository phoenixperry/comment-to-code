// Runs the Apps Script self-tests under Node: `npm test` from test/.
const path = require('path');
const crypto = require('crypto');
const { DOMParser } = require('@xmldom/xmldom');

const gs = (f) => require(path.join(__dirname, '..', 'apps-script', f));
const { parseComment, commentToRows, decodeEntities } = gs('Parse.gs');
Object.assign(global, gs('Parse.gs'), gs('Docx.gs'));
const { extractDocxComments } = gs('Docx.gs');
const { testCases_ } = gs('Tests.gs');

const XmldomAdapter = {
  parse: (s) => new DOMParser().parseFromString(s, 'text/xml').documentElement,
  name: (el) => el.localName,
  attr: (el, local) => {
    for (const a of Array.from(el.attributes)) if (a.localName === local) return a.value;
    return null;
  },
  kids: (el) => Array.from(el.childNodes).filter((n) => n.nodeType === 1),
  text: (el) => el.textContent,
};

const cases = testCases_(parseComment, (parts, hash) => extractDocxComments(parts, XmldomAdapter, hash), commentToRows);

// Real hash path too, to catch key collisions between comments.
const sha = (s) => crypto.createHash('sha1').update(s).digest('hex').slice(0, 12);
const keys = extractDocxComments(gs('Tests.gs').DOCX_FIXTURE, XmldomAdapter, sha).map((c) => c.localKey);
cases.push({ name: 'docx: sha1 keys unique', ok: new Set(keys).size === keys.length, got: keys, want: 'unique' });

let failed = 0;
for (const c of cases) {
  if (c.ok) console.log('✓ ' + c.name);
  else {
    failed++;
    console.log('✗ ' + c.name + '\n  got  ' + JSON.stringify(c.got) + '\n  want ' + JSON.stringify(c.want));
  }
}
console.log(`\n${cases.length - failed}/${cases.length} passed`);
process.exit(failed ? 1 : 0);
