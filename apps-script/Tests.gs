/**
 * Self-tests for the pure parsing code. Run from Coding → Run self-tests,
 * or locally with `node test/run.js` (same cases, xmldom instead of XmlService).
 */

var W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

var DOCX_FIXTURE = {
  'word/document.xml':
    '<w:document ' + W + '><w:body>' +
    '<w:p><w:r><w:t xml:space="preserve">Interviewer: why the studio? </w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">P03: It\'s </w:t></w:r>' +
    '<w:commentRangeStart w:id="0"/><w:r><w:t>the only place</w:t></w:r>' +
    '<w:r><w:delText>deleted words</w:delText></w:r>' +
    '<w:r><w:t xml:space="preserve"> I\'m allowed to fail.</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Nobody grades it.</w:t></w:r><w:commentRangeEnd w:id="0"/>' +
    '<w:commentRangeStart w:id="2"/><w:r><w:t>Money is always tight.</w:t></w:r><w:commentRangeEnd w:id="2"/></w:p>' +
    '</w:body></w:document>',
  'word/comments.xml':
    '<w:comments ' + W + '>' +
    '<w:comment w:id="0" w:author="Phoenix" w:date="2026-09-20T10:15:00Z">' +
    '<w:p w14:paraId="0A000001"><w:r><w:t>#belonging #fail-safe</w:t></w:r></w:p>' +
    '<w:p w14:paraId="0A000002"><w:r><w:t>Studio as refuge?</w:t></w:r></w:p></w:comment>' +
    '<w:comment w:id="1" w:author="Sam" w:date="2026-09-21T09:00:00Z">' +
    '<w:p w14:paraId="0A000003"><w:r><w:t>Agree, links to P02</w:t></w:r></w:p></w:comment>' +
    '<w:comment w:id="2" w:author="Phoenix" w:date="2026-09-20T10:20:00Z">' +
    '<w:p w14:paraId="0A000004"><w:r><w:t>#precarity</w:t></w:r></w:p></w:comment>' +
    '</w:comments>',
  'word/commentsExtended.xml':
    '<w15:commentsEx ' + W + '>' +
    '<w15:commentEx w15:paraId="0A000002" w15:done="0"/>' +
    '<w15:commentEx w15:paraId="0A000003" w15:paraIdParent="0A000002" w15:done="0"/>' +
    '<w15:commentEx w15:paraId="0A000004" w15:done="1"/>' +
    '</w15:commentsEx>'
};

function testCases_(parseComment, extractDocx, commentToRows) {
  var eq = function (a, b) { return JSON.stringify(a) === JSON.stringify(b); };
  var cases = [];
  function t(name, got, want) { cases.push({ name: name, ok: eq(got, want), got: got, want: want }); }

  t('two tags + memo', parseComment('#belonging #precarity\nShe frames the studio as refuge.'),
    { codes: ['belonging', 'precarity'], memo: 'She frames the studio as refuge.' });
  t('no tags → uncoded', parseComment('Just noticing the pause here.'),
    { codes: [''], memo: 'Just noticing the pause here.' });
  t('tag mid-sentence, case folded, deduped', parseComment('This is #Care again, #care'),
    { codes: ['care'], memo: 'This is again,' });
  t('unicode + hyphen', parseComment('#pertenencia #fail-safe #cafés'),
    { codes: ['pertenencia', 'fail-safe', 'cafés'], memo: '' });
  t('URL fragments ignored', parseComment('see https://ex.com/page#section and &#39;'),
    { codes: [''], memo: 'see https://ex.com/page#section and &#39;' });
  t('bare # and numbers-first ignored', parseComment('# heading and #1 and #_x'),
    { codes: [''], memo: '# heading and #1 and #_x' });
  t('trailing punctuation', parseComment('(#identity-) #voice.'),
    { codes: ['identity', 'voice'], memo: '() .' });

  if (typeof decodeEntities === 'function') {
    t('entities decoded', decodeEntities('you&#39;re &amp; &quot;us&quot; &#x2014; &bogus;'),
      'you\'re & "us" \u2014 &bogus;');
  }

  var rows = commentToRows({
    key: 'AAA', fileId: 'F1', fileName: 'P03 "final".docx', link: 'https://x/d?disco=AAA',
    excerpt: '-yeah, it was', content: '#care #voice\nHesitates.', author: 'Phoenix',
    created: '2026-09-20T10:00:00Z', modified: '2026-09-20T10:05:00Z', resolved: false, deleted: false,
    replies: [{ content: 'Me too', author: 'Sam', created: '2026-09-21T08:00:00Z', deleted: false },
              { content: 'gone', author: 'Sam', created: '2026-09-21T08:01:00Z', deleted: true }],
    sourceType: 'gdoc'
  });
  t('rows: one per code', rows.map(function (r) { return r[0]; }), ['AAA#care', 'AAA#voice']);
  t('rows: memo includes live replies only', rows[0][6], 'Hesitates.\n↳ Sam (2026-09-21): Me too');
  t('rows: hyperlink escapes quotes', rows[0][3], '=HYPERLINK("https://x/d?disco=AAA","P03 ""final"".docx")');

  // Subcodes and board recodes.
  t('subcode tag', parseComment('#failure/ritual #care/').codes, ['failure/ritual', 'care']);
  t('recode reply swaps code in place', applyRecodes(['privacy', 'mood'], [{ content: '↻ #mood → #atmosphere' }]),
    ['privacy', 'atmosphere']);
  t('recode reply codes an uncoded comment', applyRecodes([''], [{ content: '↻ # → #care' }]), ['care']);
  t('deleted recode reply ignored', applyRecodes(['a'], [{ content: '↻ #a → #b', deleted: true }]), ['a']);
  t('recode replies kept out of memo', commentToRows({
    key: 'K', fileId: 'F', fileName: 'f', content: '#a #b', sourceType: 'gdoc',
    replies: [{ content: '↻ #b → #c' }, { content: 'hi', author: 'S' }]
  }).map(function (r) { return [r[0], r[6]]; }), [['K#a', '↳ S: hi'], ['K#c', '↳ S: hi']]);
  t('replace tag keeps the rest', replaceTagInText('#privacy #mood\nnote', 'mood', 'atmosphere'),
    '#privacy #atmosphere\nnote');
  t('replace tag respects boundaries', replaceTagInText('#moody', 'mood', 'x'), null);
  t('replace hyphen/subcode tag', replaceTagInText('#fail-safe #a/b', 'a/b', 'a/c'), '#fail-safe #a/c');
  t('add tag to uncoded text', replaceTagInText('note', '', 'care'), '#care note');
  var rx = rewriteDocxCommentTag(DOCX_FIXTURE['word/comments.xml'], '0', 'fail-safe', 'permission-to-fail');
  t('docx rewrite edits one comment only', rx && rx.replace('#belonging #permission-to-fail', '#belonging #fail-safe')
    .replace('<w:t xml:space="preserve">#belonging #fail-safe', '<w:t>#belonging #fail-safe'),
    DOCX_FIXTURE['word/comments.xml']);
  t('docx rewrite: missing tag → null', rewriteDocxCommentTag(DOCX_FIXTURE['word/comments.xml'], '2', 'nope', 'x'), null);

  var hash = function (s) { return 'h(' + s + ')'; };
  var dx = extractDocx(DOCX_FIXTURE, hash);
  t('docx: two top-level comments', dx.length, 2);
  t('docx: excerpt spans paragraphs, skips delText', dx[0] && dx[0].excerpt,
    'the only place I\'m allowed to fail.\nNobody grades it.');
  t('docx: multi-paragraph comment text', dx[0] && dx[0].content, '#belonging #fail-safe\nStudio as refuge?');
  t('docx: reply attached', dx[0] && dx[0].replies,
    [{ content: 'Agree, links to P02', author: 'Sam', created: '2026-09-21T09:00:00Z', deleted: false }]);
  t('docx: key from author+date', dx[0] && dx[0].localKey, 'h(Phoenix|2026-09-20T10:15:00Z)');
  t('docx: done → resolved', dx[1] && [dx[1].resolved, dx[1].excerpt], [true, 'Money is always tight.']);
  return cases;
}

/** Apps Script runner. */
function runTests() {
  var cases = testCases_(parseComment, function (parts, hash) {
    return extractDocxComments(parts, XmlServiceAdapter, hash);
  }, commentToRows);
  var failed = cases.filter(function (c) { return !c.ok; });
  var msg = (cases.length - failed.length) + '/' + cases.length + ' passed' +
    failed.map(function (c) {
      return '\n✗ ' + c.name + '\n  got  ' + JSON.stringify(c.got) + '\n  want ' + JSON.stringify(c.want);
    }).join('');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from editor */ }
  return failed.length === 0;
}

if (typeof module !== 'undefined') {
  module.exports = { testCases_: testCases_, DOCX_FIXTURE: DOCX_FIXTURE };
}
