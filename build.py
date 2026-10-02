# -*- coding: utf-8 -*-
"""يبني ملف الوورد من ملف المصدر (بلا python-docx؛ OOXML مباشرة عبر zipfile).
الاستعمال:  python3 build.py source.txt out.docx [--real-footnotes]
- #3 #4 #5 عناوين، وما عداها فقرات متن.
- الحواشي بين < > في المصدر: تبقى في المتن كما هي (الأصل)، أو تحول إلى حواشي وورد حقيقية مع --real-footnotes.
- @@( )@@ ملاحظات تحريرية تظهر مظللة بالأصفر (لا يسلم ملف فيها ملاحظة).
"""
import re, sys, zipfile, datetime
from xml.sax.saxutils import escape

NBSP = ' '
HINDI = str.maketrans('0123456789', '٠١٢٣٤٥٦٧٨٩')
PHRASES = ['صلى الله عليه وسلم', 'رضي الله عنه', 'رضي الله عنهما', 'رحمه الله', 'عليه السلام',
           'سبحانه وتعالى', 'عز وجل']

def normalize(s):
    s = s.replace(NBSP, ' ')                 # تنظيف المسافات غير المنقسمة الواردة من النسخ
    s = s.replace('—', '-').replace('–', '-')  # الشرطة الطويلة تبدل شرطة
    s = re.sub(r'([ء-ي])اً', r'\1ًا', s)  # تنوين النصب على الحرف قبل الألف
    s = re.sub(r'(\d+)/ (\d)', lambda m: m.group(1) + '/' + NBSP + m.group(2), s)   # (٣/ ٦٦)
    s = re.sub(r'\(ص: ', '(ص:' + NBSP, s)
    s = re.sub(r'رقم (\d)', lambda m: 'رقم' + NBSP + m.group(1), s)
    for p in PHRASES:
        s = s.replace(p, p.replace(' ', NBSP))
    s = s.translate(HINDI)                   # كل الأرقام هندية
    return s

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
FONT = '<w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic" w:cs="Traditional Arabic" w:eastAsia="Traditional Arabic"/>'

def run(text, hl=False, bold=False, style=None, size=None):
    rpr = FONT
    if style: rpr = '<w:rStyle w:val="%s"/>' % style + rpr
    if bold: rpr += '<w:b/><w:bCs/>'
    if size: rpr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size, size)
    if hl: rpr += '<w:highlight w:val="yellow"/>'
    rpr += '<w:rtl/>'
    return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(text))

def split_notes(text):
    """يقسم النص على الملاحظات الصفراء @@( )@@"""
    out = []
    for part in re.split(r'(@@\(.*?\)@@)', text, flags=re.S):
        if part.startswith('@@(') and part.endswith(')@@'):
            out.append((part[3:-3], True))
        elif part:
            out.append((part, False))
    return out

class Builder:
    def __init__(self, real_footnotes=False):
        self.real = real_footnotes
        self.footnotes = []
    def para_runs(self, text):
        runs = []
        for seg, hl in split_notes(text):
            if self.real and not hl:
                for piece in re.split(r'(<[^>]*>)', seg):
                    if piece.startswith('<') and piece.endswith('>'):
                        self.footnotes.append(piece[1:-1])
                        fid = len(self.footnotes)
                        runs.append('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/>%s<w:rtl/></w:rPr><w:footnoteReference w:id="%d"/></w:r>' % (FONT, fid + 1))
                    elif piece:
                        runs.append(run(piece))
            else:
                runs.append(run(seg, hl=hl))
        return ''.join(runs)
    def paragraph(self, text):
        return ('<w:p><w:pPr><w:pStyle w:val="BodyText1"/></w:pPr>%s</w:p>' % self.para_runs(text))
    def heading(self, level, text):
        return ('<w:p><w:pPr><w:pStyle w:val="Heading%d"/></w:pPr>%s</w:p>' % (level, run(text, bold=True)))

def build_document(src_lines, real):
    b = Builder(real)
    body = []
    for line in src_lines:
        line = line.rstrip('\n')
        if not line.strip():
            continue
        line = normalize(line)
        m = re.match(r'^#([1-5]) (.*)$', line)
        if m:
            body.append(b.heading(int(m.group(1)), m.group(2)))
        else:
            body.append(b.paragraph(line))
    sect = ('<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/>'
            '<w:footnotePr><w:numFmt w:val="hindiNumbers"/></w:footnotePr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="709" w:footer="709" w:gutter="0"/>'
            '<w:pgNumType w:fmt="hindiNumbers"/><w:bidi/></w:sectPr>')
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s%s</w:body></w:document>' % (W, ''.join(body), sect)
    return doc, b.footnotes

def footnotes_xml(notes):
    items = ['<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
             '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>']
    for i, n in enumerate(notes, 1):
        items.append('<w:footnote w:id="%d"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
                     '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/>%s<w:rtl/></w:rPr><w:footnoteRef/></w:r>'
                     '<w:r><w:rPr>%s<w:sz w:val="28"/><w:szCs w:val="28"/><w:rtl/></w:rPr><w:t xml:space="preserve"> %s</w:t></w:r></w:p></w:footnote>'
                     % (i + 1, FONT, FONT, escape(n)))
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes %s>%s</w:footnotes>' % (W, ''.join(items))

def styles_xml():
    def hstyle(n, size):
        return ('<w:style w:type="paragraph" w:styleId="Heading%d"><w:name w:val="heading %d"/><w:basedOn w:val="Normal"/><w:next w:val="BodyText1"/><w:qFormat/>'
                '<w:pPr><w:keepNext/><w:bidi/><w:spacing w:before="240" w:after="160" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/><w:outlineLvl w:val="%d"/></w:pPr>'
                '<w:rPr>%s<w:b/><w:bCs/><w:sz w:val="%d"/><w:szCs w:val="%d"/><w:rtl/></w:rPr></w:style>' % (n, n, n - 1, FONT, size, size))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>'
            '<w:docDefaults><w:rPrDefault><w:rPr>%s<w:sz w:val="36"/><w:szCs w:val="36"/><w:lang w:val="ar-SA" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:bidi/><w:spacing w:after="160" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
            '<w:style w:type="paragraph" w:styleId="BodyText1"><w:name w:val="Body Text 1"/><w:basedOn w:val="Normal"/><w:qFormat/>'
            '<w:pPr><w:bidi/><w:spacing w:after="160" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="283"/><w:jc w:val="both"/></w:pPr>'
            '<w:rPr>%s<w:sz w:val="36"/><w:szCs w:val="36"/><w:rtl/></w:rPr></w:style>'
            '%s%s%s%s%s'
            '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
            '<w:pPr><w:bidi/><w:spacing w:after="40" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr><w:rPr>%s<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
            '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>'
            '</w:styles>') % (W, FONT, FONT, hstyle(1, 44), hstyle(2, 40), hstyle(3, 44), hstyle(4, 40), hstyle(5, 36), FONT)

def footer_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr %s><w:p><w:pPr><w:pStyle w:val="Footer"/><w:bidi/><w:jc w:val="center"/></w:pPr>'
            '<w:r><w:rPr>%s<w:rtl/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:rPr>%s<w:rtl/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:rPr>%s<w:rtl/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:rPr>%s<w:rtl/></w:rPr><w:t>١</w:t></w:r>'
            '<w:r><w:rPr>%s<w:rtl/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>') % (W, FONT, FONT, FONT, FONT, FONT)

def write_docx(path, src, real, title):
    doc, notes = build_document(src, real)
    now = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
          + ('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' if real else '') +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
             + ('<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' if real else '') +
             '</Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>'
                '<w:characterSpacingControl w:val="doNotCompress"/>'
                + ('<w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>' if real else '') +
                '<w:themeFontLang w:val="en-US" w:bidi="ar-SA"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>') % W
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>عبدالله العزب</dc:creator><cp:lastModifiedBy>عبدالله العزب</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified></cp:coreProperties>') % (escape(title), now, now)
    app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>build.py</Application></Properties>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', doc)
        z.writestr('word/_rels/document.xml.rels', drels)
        z.writestr('word/styles.xml', styles_xml())
        z.writestr('word/settings.xml', settings)
        z.writestr('word/footer1.xml', footer_xml())
        if real:
            z.writestr('word/footnotes.xml', footnotes_xml(notes))
        z.writestr('docProps/core.xml', core)
        z.writestr('docProps/app.xml', app)
    return len(notes)

if __name__ == '__main__':
    src_path, out_path = sys.argv[1], sys.argv[2]
    real = '--real-footnotes' in sys.argv
    title = 'التحسين والتقبيح العقليان'
    n = write_docx(out_path, open(src_path, encoding='utf-8').read().split('\n'), real, title)
    print('built', out_path, 'real footnotes:', n if real else 'no (markers kept inline)')
