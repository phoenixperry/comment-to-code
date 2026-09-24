"""Build sample-data/P99_sample_interview.docx: a fictional transcript with
Word-desktop comments (incl. a reply and a resolved one) for end-to-end tests."""
import zipfile, pathlib

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
     'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')

def p(*runs): return '<w:p>' + ''.join(runs) + '</w:p>'
def t(s): return f'<w:r><w:t xml:space="preserve">{s}</w:t></w:r>'
def start(i): return f'<w:commentRangeStart w:id="{i}"/>'
def end(i): return f'<w:commentRangeEnd w:id="{i}"/><w:r><w:commentReference w:id="{i}"/></w:r>'

doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + ''.join([
    p(t('Interviewer: What does the studio mean to you?')),
    p(t('P99: Honestly? '), start(0), t("It's the only place I'm allowed to fail."), end(0),
      t(' Nobody grades it.')),
    p(t('Interviewer: And outside the studio?')),
    p(t('P99: '), start(1), t('Rent goes up every year and the commissions don\'t.'), end(1),
      t(' So I teach, which I like, mostly.')),
    p(t('P99: '), start(3), t('I think of the other makers here as family, weirdly.'), end(3)),
]) + '<w:sectPr/></w:body></w:document>'

def comment(i, author, date, para_id, text):
    return (f'<w:comment w:id="{i}" w:author="{author}" w:date="{date}" w:initials="{author[0]}">'
            f'<w:p w14:paraId="{para_id}"><w:r><w:annotationRef/></w:r>{t(text)}</w:p></w:comment>')

comments = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {W}>' + ''.join([
    comment(0, 'Researcher', '2026-09-20T10:15:00Z', '1A000001', '#permission-to-fail #refuge Studio as a place outside assessment.'),
    comment(1, 'Researcher', '2026-09-20T10:18:00Z', '1A000002', '#precarity #teaching'),
    comment(2, 'Second coder', '2026-09-21T09:00:00Z', '1A000003', 'Teaching as ballast, not just income?'),
    comment(3, 'Researcher', '2026-09-20T10:22:00Z', '1A000004', '#belonging'),
]) + '</w:comments>'

ext = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx {W}>' + ''.join([
    '<w15:commentEx w15:paraId="1A000001" w15:done="0"/>',
    '<w15:commentEx w15:paraId="1A000002" w15:done="0"/>',
    '<w15:commentEx w15:paraId="1A000003" w15:paraIdParent="1A000002" w15:done="0"/>',
    '<w15:commentEx w15:paraId="1A000004" w15:done="1"/>',
]) + '</w15:commentsEx>'

ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>'
      '</Types>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>'
            '</Relationships>')

out = pathlib.Path(__file__).with_name('P99_sample_interview.docx')
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct)
    z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc)
    z.writestr('word/_rels/document.xml.rels', doc_rels)
    z.writestr('word/comments.xml', comments)
    z.writestr('word/commentsExtended.xml', ext)
print('wrote', out)
