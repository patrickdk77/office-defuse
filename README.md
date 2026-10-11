# office-defuse

office-defuse reads Word, Excel, PowerPoint and Visio files, in their ZIP formats and in the compound file formats before them, OpenDocument packages, Outlook messages, Publisher and Project files, and RTF. It reports VBA, WordBasic and Excel 4.0 macros, LibreOffice Basic and scripts, DDE, fields and events that run code, templates and objects loaded from outside the file, embedded OLE objects, Equation Editor objects, attachments, ActiveX controls and unsafe links. It scores the risk from 0 to 100.

This version inspects and changes nothing.

It has no runtime dependencies and keeps memory bounded. It reads ZIP archives, compound files and XML strictly, and refuses anything two readers could take in different ways. RTF it reads the way Word does, quirks included.

```
npm install @patrickdk77/office-defuse
```

Node 18 or later.

## Quick start

```js
const { inspectOffice } = require('@patrickdk77/office-defuse');

const report = await inspectOffice(bytes);
report.status;   // 'clean', 'strippable' or 'rejected'
report.risk;     // 'NONE', 'LOW', 'MEDIUM', 'HIGH', 'CRITICAL' or 'UNKNOWN'
report.score;    // 0 to 100, null when the file could not be inspected
report.format;   // 'docx', 'xls', 'msg', 'rtf', ... or '' when not recognized
report.mimeType; // the MIME type of that format, or ''
report.findings; // [{ category, detail, description, action, location, data }]
```

ES modules work too: `import { inspectOffice } from '@patrickdk77/office-defuse'`.

`strippable` means at least one finding has the action `strip`, content a defused copy would leave out. `rejected` means a finding has the action `reject`, or the run stopped before it read the whole file.

### Files and S3

`inspectOfficeSource` reads through a source with `size()` and `read(offset, length)`, so a local file or an S3 object never has to sit in memory. `fileSource` and `bufferSource` are included. For S3, `read` can be a ranged GetObject.

```js
const { inspectOfficeSource, fileSource } = require('@patrickdk77/office-defuse');

await inspectOfficeSource(fileSource('upload.docx'));

const source = {
  size: async () => contentLength,
  read: async (offset, length) => getObjectRange(bucket, key, offset, offset + length - 1),
};
await inspectOfficeSource(source);
```

`inspectOfficeSource` calls the source's `close()`, if it has one, when the run ends. A `read` must return exactly the bytes asked for, or fewer only at the end of the source. Any other length fails the run with an error whose `code` is `EIO`. An error the source throws comes back unchanged, instead of a verdict on the file.

## Formats

The format comes from the content, never from the file name.

| Family | ZIP formats | Compound file and RTF formats |
|---|---|---|
| Word | .docx .docm .dotx .dotm | .doc .dot from Word 6.0 to 2003, and .rtf |
| Excel | .xlsx .xlsm .xltx .xltm .xlam .xlsb | .xls .xlt .xla from Excel 5.0 to 2003 |
| PowerPoint | .pptx .pptm .potx .potm .ppsx .ppsm .ppam .sldx .sldm | .ppt from PowerPoint 97 to 2003 |
| Visio | .vsdx .vsdm .vstx .vstm .vssx .vssm | .vsd from Visio 2003 to 2010 |
| OpenDocument | .odt .ott .ods .ots .odp .otp .odg .otg .odf .odm .odb | |
| Outlook | | .msg .oft |
| Publisher | | .pub |
| Project | | .mpp |

`[Content_Types].xml` marks an Office Open XML package. The package relationship of type `officeDocument`, transitional or strict, or Visio's `document`, names the main part, and its content type gives the format. The `mimetype` entry marks an OpenDocument package, or the manifest's entry for `/` when there is no mimetype entry.

A ZIP with both markers is rejected as PACKAGE/AMBIGUOUS_FORMAT, since Office could read one and LibreOffice the other. So is an OOXML package with two main parts, and an ODF package whose mimetype and manifest name different formats. Anything else is PACKAGE/UNRECOGNIZED.

A .xlsb gets the package checks. Its sheets are binary records, and the formulas in them go unread. Its external link parts are binary too, and office-defuse reads their link records for DDE links. A .odb gets every ODF check, and the database under `database/` goes unread.

A file that starts with the compound file signature is a compound file, MS-CFB. The streams and storages at its root name the format: `WordDocument` for Word, `Workbook` or `Book` for Excel, `PowerPoint Document`, `VisioDocument`, `__properties_version1.0` for an Outlook message, a `Quill` storage or Publisher's class id beside a `Contents` stream for Publisher, and a `Props` stream beside a storage named by a number for Project. Word's FIB says whether a document is a template, and Excel's records whether a workbook is a template or an add-in. An Outlook template has a class id of its own. A PowerPoint 97 to 2003 template or show reports its format as .ppt, and a Visio stencil or template as .vsd.

A compound file with two of these is PACKAGE/AMBIGUOUS_FORMAT. One with none is PACKAGE/UNRECOGNIZED, and so is a `WordDocument` stream that holds no Word 6.0 to 2003 FIB. A PowerPoint 4.0 file, a bare OLE object and a `vbaProject.bin` on its own all land there.

An encrypted .docx, .xlsx or .pptx is a compound file holding `EncryptionInfo` and `EncryptedPackage`. See [Encrypted files](#encrypted-files).

A file that starts with `{\rt` is RTF, as Word takes it. A compound file or RTF file in which the ZIP reader also finds an archive is PACKAGE/AMBIGUOUS_FORMAT. A file with the compound file signature whose compound file does not read, but whose ZIP archive does, is that ZIP archive with LEADING_BYTES.

## What it reports

In a ZIP package:

- VBA projects, Excel 4.0 macro sheets, and OpenDocument `Basic/`, `Scripts/` and `Dialogs/` folders. These go by content type, relationship type or folder name, and a folder counts when it holds a file that is not empty. office-defuse parses no OLE part of a package, such as `vbaProject.bin`.
- OpenDocument script event listeners that name a script URL or macro. Those that fire as the document opens or is created get a finding of their own. LibreOffice ignores a listener whose URL and macro name are both empty, and so does office-defuse. Visio ShapeSheet cells that call RUNADDON, RUNADDONWARGS, CALLTHIS, RUNMACRO, OPENFILE or DOOLEVERB.
- Links and actions that run macros: `vnd.sun.star.script:` and `macro:` targets, PowerPoint's `ppaction://macro`, Excel drawing shapes with a `macro` attribute, and form controls with an `x:FmlaMacro`. Ribbon callbacks such as `onLoad` and `onAction`. OpenDocument menu, toolbar and key settings under `Configurations2/`. LibreOffice writes an empty `Configurations2/accelerator/current.xml` into the documents it saves, and an empty file counts for nothing.
- Excel defined names that run a macro sheet: Auto_Open and Auto_Activate as MACRO/EVENT_ON_OPEN, Auto_Close and Auto_Deactivate as MACRO/EVENT. A name counts when it starts with one of these, as `Auto_Open2` does.
- PowerPoint and Impress actions that run a program or an object verb.
- DDE: Word `DDE` and `DDEAUTO` fields, Excel `ddeLink` parts and the .xlsb record for them, Excel formulas in `application|topic!item` form, OpenDocument DDE connections and `DDE()` formulas. No amount of white space before a field's name hides it, and a formula is checked whole however long it runs.
- Word fields that run a macro, `MACROBUTTON`, and fields whose type another field builds. That is the QUOTE trick that hides `DDEAUTO`.
- References that load from outside the file on their own: attached and OpenDocument templates, linked OLE objects and packages, images, frames, included files and linked sections, external workbooks, `meta:auto-reload`, and any other external relationship. A target on the machine itself is EXTERNAL/LOCAL_FILE. An empty OpenDocument reference is the document itself and names no file, so it goes unreported. A target counts as local only if its percent-decoded form is local too, because a CVE-2017-0199 sample writes its `script:` scheme as `%73cript:`.
- Hyperlinks, with pdf-defuse's checks of the target: the scheme, file and network paths, data URLs, credentials before the host, IP hosts, punycode and mixed-script hosts, and percent-encoded hosts. A link to a place inside the document goes unreported, and so does an empty link, which is the document itself.
- Embedded OLE objects and packages, OLE files nothing declares as an object, Java applets, plug-ins and ActiveX controls.
- Package structure: parts without a content type, parts no relationship reaches, files the ODF manifest does not list, an irregular ODF layout, and names a package gives twice.
- Document type declarations in XML parts, and references to external entities. See [Document types and entities](#document-types-and-entities).

The OOXML rules match elements by namespace URI and local name, compared ignoring case, so a prefix bound to the Word namespace reads as Word whatever it is called. The OpenDocument rules read `urn:oasis:names:tc:<any>:xmlns:<name>:1.<any>` as version 1.0 of that namespace, which is what LibreOffice does. Every branch of `mc:AlternateContent` gets read.

In a compound file, the first two rules run over every storage of the tree, and the others over the streams of the format:

- VBA projects: a `VBA` storage with a `dir` stream, or a `_VBA_PROJECT_CUR` or `Macros` storage with something in it.
- Embedded objects, by the streams an object's storage holds, under `ObjectPool` or anywhere else: OLE objects with their class id, ProgID and user type, the file a Packager object holds, Equation Editor 3.0 objects by class id or ProgID, ActiveX controls, and linked objects with the target their moniker names. An object's storage is part of the tree, so a macro or an object inside an object is found.
- Word: WordBasic macros in Word 6.0 and 95 files. The fields of the text, with the rules a .docx gets, and the attached template.
- Excel: Excel 4.0 macro sheets, by sheet type or by the sheet's own BOF record, and the Auto_ names above, built in or defined. External workbooks, with the paths Excel encodes written out, DDE and OLE links, hyperlinks, shapes and buttons that run a macro, and the `Ctls` stream that holds ActiveX controls.
- PowerPoint: actions that run a macro, a program or an object verb, hyperlinks, and linked sounds and movies. PowerPoint keeps the storages of objects, controls and the VBA project in its document stream, often compressed. office-defuse inflates them and gives them the rules above.
- Outlook: each attachment, as EMBEDDED/ATTACHMENT with its name, MIME type and size, and each attachment by reference with its path. An attached message's own attachments are read, 16 messages deep.
- Visio, Publisher and Project: the VBA projects and objects. Nothing in their own streams is read.

In RTF:

- Objects in `\objdata`. office-defuse decodes the OLE 1.0 object, the file a Packager object holds, and the storage of an OLE 2 object, which gets the compound file rules. An object linked with `\objlink` or `\objautlink`, or of the class OLE2Link that CVE-2017-0199 uses, is EXTERNAL/OBJECT. `\objocx` marks an ActiveX control. `\objupdate` is ACTION/OBJECT_UPDATE, since it starts the object's program as the document opens.
- Fields, with the rules a .docx gets, the `\*\template` destination, and shape hyperlinks.
- Text after the group that closes the document, as TRAILING_BYTES. White space and NULs there are what writers leave, and count for nothing.

### Encrypted files

office-defuse decrypts an OOXML package encrypted with Standard or Agile encryption, MS-OFFCRYPTO, and an .xls workbook encrypted with RC4 or RC4 CryptoAPI. It tries the `password` option first, then VelvetSweatshop, the password Excel tries on its own, so such a workbook opens without asking. The decrypted file gets every check of its format. It is read in place. Agile encryption decrypts in 4,096-byte segments and Standard in 16-byte blocks, so only the parts office-defuse reads get decrypted.

| Finding | Action | Weight | When |
|---|---|---|---|
| ENCRYPTED/DEFAULT_PASSWORD | `strip` | 10 | VelvetSweatshop opened the file. Anyone can open it, so the encryption hides its content from a scanner and from no one else |
| ENCRYPTED/USER_PASSWORD | `strip` | 0 | The password the caller gave opened the file |
| ENCRYPTED/RC4_40, RC4_128, RC4_OTHER, AES_128, AES_192, AES_256 | `info` | 0 | The cipher and key size |
| ENCRYPTED/PASSWORD_REQUIRED | `reject` | 0 | No password at hand opens the file, or an Agile spin count is above 10,000,000. Every encrypted .doc and .ppt, which office-defuse does not decrypt |
| ENCRYPTED/UNSUPPORTED_ALGORITHM | `reject` | 0 | Excel's XOR obfuscation, or a cipher or hash outside the list below |

AES-128, AES-192 and AES-256 in CBC or ECB mode, SHA-1, SHA-256, SHA-384, SHA-512 and MD5 go to Node's crypto module, and no other algorithm name from a file reaches it. RC4 is office-defuse's own code, since OpenSSL 3 keeps RC4 in its legacy provider.

Pass the password in code, never on a command line, where other users can read it.

## The containers

The ZIP reader starts from the end record and the central directory. It checks each local header against the central directory before it reads any content. It refuses these, and stops the run:

| Detail | Refused |
|---|---|
| NOT_ZIP | No end record in the last 65,557 bytes, or one whose comment runs past the end of the file |
| AMBIGUOUS_END | An end record whose comment holds a second end signature |
| MULTI_DISK | Any disk number other than 0 |
| DIRECTORY_MISMATCH | An entry count or directory size that disagrees with the central directory, or a central directory that does not end where the end record begins. Data prepended without fixing the offsets lands here |
| OUT_OF_RANGE | A central directory or local header outside the file or its entries, or a 64-bit value larger than a JavaScript number holds exactly |
| OVERLAP | Two entries that share bytes, or an entry that runs into the central directory |
| HEADER_MISMATCH | A local header that is missing, or whose name, method, encryption, data descriptor or UTF-8 flags, CRC or sizes differ from the central directory, and a data descriptor that differs |
| DUPLICATE_NAME | Two names equal after percent-decoding, Unicode NFC and lowercasing |
| UNSAFE_NAME | An empty name, a name over 1024 bytes, a leading `/`, a drive letter, a backslash, a control character, an empty, `.` or `..` segment, or a directory entry with content |
| NAME_ENCODING | A non-ASCII name without the UTF-8 flag, invalid UTF-8, or an Info-ZIP Unicode path field that names another entry |
| ENCRYPTED_ENTRY | Any of the encryption flags |
| UNSUPPORTED_METHOD | Any compression method but stored and deflate |
| ZIP64_MISMATCH | A field at its maximum with no ZIP64 value behind it, a ZIP64 extra field that holds other fields than its header defers to it, a local ZIP64 field that disagrees with its header, an end record that disagrees with the ZIP64 end record, or a locator that does not point at one |
| BAD_EXTRA_FIELD | An extra field block that does not split exactly into records, or a record id given twice |
| DATA_ERROR | Data that does not inflate to its declared size and CRC, deflate data that is corrupt, ends early or leaves compressed bytes unused, or a stored entry with two different sizes |

It reports three more without stopping. LEADING_BYTES is data before the first entry. HIDDEN_DATA is bytes between entries, or before the central directory, that no entry covers. TRAILING_BYTES is bytes after the end record that its comment does not cover. Each is `strip`.

The reader inflates every entry once, start to end, and checks it against its CRC and declared size. Inflation stops as soon as the output passes the declared size.

A writer that streams an entry sets the data descriptor flag and may leave the local header's CRC and sizes unset, as zero, or as ones in a local ZIP64 field. The data descriptor after the data then gives them, and they must match the central directory.

The XML tokenizer reads each part in pieces. It refuses these, and stops the run:

- XML/ENCODING: anything but UTF-8 or UTF-16, the two encodings OPC allows, an encoding declaration that disagrees with the bytes, and bytes that are not valid in their encoding.
- XML/MALFORMED: markup that breaks XML 1.0 or its namespaces. That covers an undeclared prefix, a duplicate attribute, a mismatched end tag, an element or text outside the root, a malformed name, reference or declaration, and characters XML 1.0 forbids. It also covers what XML 1.0 refuses about entities: a reference to an entity nothing declares, a reference to an external entity in an attribute value, entity text that refers to itself or leaves an element open, and a markup declaration outside the document type declaration.

References get expanded and line ends normalized as the XML specification says, so `DD&#69;AUTO` reads as `DDEAUTO`.

The `Basic/`, `Scripts/`, `Dialogs/` and `Configurations2/` folders of an OpenDocument package are reported by name, as a whole. The files in them are inflated and checked, not parsed. In an OOXML package, the parts read as XML are `[Content_Types].xml`, every `.rels` part, and every part whose content type ends in `+xml`, is `application/xml` or `text/xml`, or is the VML drawing type. In an OpenDocument package they are the other entries whose names end in `.xml`.

PowerPoint and Excel write VML drawings that are not well-formed XML. They put `<![if gte mso 9]>` and `<![endif]>` around content, and Excel 2007 leaves `<br>` unclosed in a button's text. office-defuse reads a part of the VML drawing type in that dialect. It skips the markers, so the content on both sides of them reaches the rules, and it takes `<br>` as an empty element. Apart from those two, it reads the part as strictly as any other.

### Compound files

The compound file reader follows MS-CFB, and where MS-CFB leaves room, it accepts what olefile and LibreOffice accept. It reads the header, the FAT through the DIFAT, the directory, the mini FAT and the mini stream, and records which stream owns each sector, so no sector serves two. It follows a stream's chain only as far as the stream's size needs, as olefile and LibreOffice do, so a chain that runs on past that does no harm. It refuses these, and stops the run:

| Detail | Refused |
|---|---|
| BAD_HEADER | A file shorter than its header, a version and sector size other than 3 with 512 bytes or 4 with 4,096, mini sectors other than 64 bytes, a mini stream cutoff other than 4,096, a byte order other than little-endian, FAT or DIFAT sector counts that disagree with the sectors listed, a mini FAT too short for the mini stream, or an empty directory |
| OUT_OF_RANGE | A sector or a stream outside the file, a directory entry that names one past the directory, or a stream size larger than a JavaScript number holds exactly |
| OVERLAP | A sector or mini sector that two streams take, or a stream and the FAT, DIFAT, directory or mini FAT |
| CHAIN_LOOP | A FAT, DIFAT or mini FAT chain that comes back to a sector it took already |
| BROKEN_CHAIN | A chain that runs into a free or reserved sector before its stream ends, or a DIFAT chain that does not end after its last sector |
| DATA_ERROR | A chain that ends before its stream does |
| DIRECTORY_LOOP | A directory entry the tree reaches twice |
| BAD_ENTRY | An entry in the tree of an unknown or unused type, a root entry that is not the first, a name length or terminator that does not fit, a NUL in a name, or a stream with children |
| UNSAFE_NAME | A name holding `/`, `\`, `:` or `!`, which MS-CFB forbids |
| DUPLICATE_NAME | Two entries of one storage whose names are equal ignoring case |

It reports two more without stopping, each `strip`. HIDDEN_DATA counts the bytes that no stream covers and that are not zero: sectors no chain takes, the rest of each stream's last sector or mini sector, mini sectors no stream takes, directory entries the tree does not reach, and the padding of a version 4 header. Zeros are padding and count for nothing. TRAILING_BYTES is data past the last sector the FAT describes. HIDDEN_DATA is common in real files. LibreOffice, for one, leaves an old copy of a directory sector in a free sector when it saves.

`limits.entries` counts directory entries, `limits.decompressedBytes` bounds the size of one stream, and `limits.totalDecompressedBytes` the sizes of all streams together, the way they count ZIP entries. The reader holds the FAT, the mini FAT, the directory and the owner of each sector, and reads streams in ranges as the rules ask for them.

### RTF

The RTF reader takes the file in 256 KiB pieces. It holds the destinations open at the point it has reached, the short text it collects, such as a class name or a template path, and the data of each object it decodes. It reads the way Word reads:

- A control word's name and its number each end after 250 characters.
- A destination that `\*` marks and that office-defuse does not know is skipped with its group, and an unknown control word is ignored. Every destination of RTF 1.9.1 is known, so a decoy group such as `{\listtable 4142}` inside an object's data adds no hex digits.
- In an object's data, anything but hex digits drops out, a digit left unpaired before `\'hh` is dropped, and `\bin` data goes in as its bytes.
- Inside an object's data, a group that `\*` marks and that opens with a control word that is no destination, such as `{\*\par574 0}`, still holds hex digits for rtfobj, and Word opens files that hide an object's data that way. When the data reads differently with and without those groups, both readings get the object rules, and the findings from the second have `reading` in their data.
- A file can end inside groups, and the reader closes them. Text after the group that closes the document is still read, as rtfobj reads it.

`limits.rtfDepth`, 256 by default, bounds group nesting and the destinations open at once. Past it the run stops with LIMIT/RTF_DEPTH. One object's decoded data counts against `limits.decompressedBytes`, and all of them together against `limits.totalDecompressedBytes`.

### Document types and entities

office-defuse never loads an external DTD and never fetches an external entity. A document type declaration does not stop the run.

| Finding | Action | Weight | What it is |
|---|---|---|---|
| XML/DOCTYPE | `info` | 0 | A declaration with a name and maybe an external identifier, and no internal subset. OpenOffice.org 2.x and StarOffice 8 wrote one into their manifests, and OpenOffice.org 2.x and 3.x into their formula objects |
| XML/INTERNAL_SUBSET | `strip` | 20 | A declaration with an internal subset, the declarations between square brackets, which can declare entities and attribute defaults |
| XML/EXTERNAL_ENTITY | `strip` | 40 | A reference to an external entity in element text |

The tokenizer expands internal entities where the part refers to them, in element text and in attribute values. The rules see the expanded text, so a relationship target, a hyperlink or a field code built from entities gets checked as its real value. Markup in entity text becomes elements.

A reference to an external entity in text stands for content office-defuse cannot read. Findings show it as `&name;`, and a rule that needs the value treats it as unknown. A hyperlink target holding one is LINK/RELATIVE, which is what pdf-defuse makes of a URI it cannot resolve. A target that loads from outside never counts as a local file. A Word field whose name holds one is FIELD/COMPUTED, and an Excel formula that holds one is DDE/FORMULA, since the unknown text could make it a DDE call. XML 1.0 forbids a reference to an external entity in an attribute value, and xmllint and expat refuse one, so office-defuse refuses it as XML/MALFORMED.

The internal subset is read the way XML 1.0 has a processor read it when that processor reads internal parameter entities and no external ones. The first declaration of a name binds. A parameter entity reference that is not read, because the entity is external or undeclared, makes the tokenizer skip the entity and attribute-list declarations after it, unless the XML declaration says `standalone="yes"`. Conditional sections and parameter entity references inside a declaration are refused. libxml2 reads further. It keeps reading declarations after a parameter entity it does not load, and it takes a parameter entity reference inside a declaration that came from another parameter entity's text.

An element that leaves out an attribute gets the default its attribute-list declaration gives, namespace declarations included. An attribute declared with a type other than CDATA has its spaces collapsed. expat does both. libxml2 does both for namespace declarations, and applies other defaults only with XML_PARSE_DTDATTR.

Three limits bound what a document type can cost: entity references nested 16 deep, 100,000 expansions, and 16 Mi characters of expanded text. The last two count over the whole run, and each attribute default added to an element counts as an expansion. Past one of them the run stops with LIMIT/XML_ENTITY_DEPTH, LIMIT/XML_ENTITY_EXPANSIONS or LIMIT/XML_ENTITY_SIZE. The billion laughs stops at the expansion limit.

### When a run stops

A run that stops rejects the file with a null score and the UNKNOWN level, whatever `actionOverrides` say about the finding that stopped it. The container refusals, a limit, an unrecognized or ambiguous format, an encrypted file that no password at hand opens or whose algorithm office-defuse does not decrypt, and an unexpected error, PACKAGE/UNREADABLE, all stop a run. The findings made before the stop stay in the report.

## Findings

Every finding has a `category` and a `detail`, plain strings such as `MACRO` and `VBA_PROJECT`, and a `description` that is safe to show the uploader. A description that quotes a value from the file cuts it to 60 characters and shows control and format characters as `?`. `data` keeps a value from the file up to 200 characters, or 2,000 for a target, and `location` names the part, the relationship, the folder, the stream or storage, or the RTF object's offset.

```js
{ category: 'EXTERNAL', detail: 'TEMPLATE', description: 'The document loads its template from https://templates.example.net/report.dotm', action: 'strip', location: 'word/_rels/settings.xml.rels, rIdTpl', data: { target: '...', type: '...' } }
```

Past 200 findings of one kind, one more finding with the location `additional occurrences` counts the rest in `data.count`.

`action` is `reject`, `strip` or `info`. Each rule decides its own action, and `info` only reports. The enums are exported as `OfficeCategory` and `OfficeDetail`, and `allFindingSpecs()` lists every pair with its default action, description and weight. `actionOverrides` changes the action by category, or by category and detail:

```js
await inspectOffice(bytes, { actionOverrides: [{ category: 'MACRO', action: 'reject' }] });
```

## Scoring

The method is pdf-defuse's. Each kind of finding has a weight: a program run by a click is 80, VBA is 70, a safe link is 0. The highest weight counts in full, every other kind adds a quarter of its weight, and repeats of one kind count once.

Three combinations add more. Code here means macros or DDE. Code or a run-program action together with a template, object or include from outside adds 20. Code with an embedded object or ActiveX control adds 15. Code with leading, hidden or trailing container data adds 10. The total is capped at 100.

The bands are NONE at 0, LOW from 1, MEDIUM from 25, HIGH from 50 and CRITICAL from 80. `scoreWeights` and `scoreBands` change them. A finding of a kind office-defuse does not make, such as one from another package, scores at its `weight`, or 0 without one.

## Options

| Option | Default | Purpose |
|---|---|---|
| `limits.fileSize` | none | Bytes in the upload, checked before any read |
| `limits.entries` | 10,000 | Entries in the central directory, or in a compound file's directory |
| `limits.decompressedBytes` | 512 MiB | The declared size of any one entry or stream, one RTF object's decoded data, or one compressed PowerPoint storage |
| `limits.totalDecompressedBytes` | 2 GiB | The declared sizes of all entries, or of all streams, together, and all RTF objects' decoded data |
| `limits.xmlDepth` | 256 | Element nesting in one XML part |
| `limits.xmlTokenSize` | 16 Mi | Characters in one tag, comment, processing instruction or CDATA section. A document type declaration counts as one. Text goes through in pieces and has no limit |
| `limits.xmlEntityDepth` | 16 | Entity references nested inside entity text |
| `limits.xmlEntityExpansions` | 100,000 | Entity references expanded, and attribute defaults added, in the whole run |
| `limits.xmlEntitySize` | 16 Mi | Characters of entity text and attribute defaults expanded in the whole run |
| `limits.rtfDepth` | 256 | Group nesting in an RTF file, and destinations open at once |
| `limits.timeMs` | none | The whole run |
| `password` | none | Tried first on an encrypted OOXML package or .xls workbook, before VelvetSweatshop |
| `actionOverrides` | none | Change the action for a category, or for a category and detail |
| `scoreWeights`, `scoreBands` | built in | Change the scoring |

The nine limits with a default bound the memory or work any input can cost, so they apply unless you set them. `Infinity` turns one off. A limit that is reached stops the run with a LIMIT finding.

Besides one XML token, the stack of open elements and the declarations of the part's internal subset, office-defuse holds the central directory's table of entries, which parts each part's relationships reach, and, while it reads a relationships part, the Ids that part has given so far, about 100 bytes each, so it can find one given twice. Content types, manifest paths and object references that name no entry of the archive are not kept. Everything else streams.

## The defuse family

office-defuse is one of a family of packages that each handle one kind of file, beside [pdf-defuse](https://github.com/patrickdk77/pdf-defuse). It declares the family's shared types with the same shapes as pdf-defuse: `DefuseFinding`, `ByteSource`, `ByteSink`, `PluginContext`, `ContainedFile`, `ContainedFileResult` and `ContainedFilePlugin`. It does not depend on pdf-defuse, and TypeScript matches the types by shape.

A finding's `category` and `detail` are plain strings, and `weight`, when set, says how much it adds to a score. The time limit's error carries `code: 'DEFUSE_LIMIT'` and `limit: 'time'`, the code every package of the family passes up the tree of files. `inspectOffice` turns it into LIMIT/TIME.

## What it does not check

- The OLE parts of a ZIP package, such as `vbaProject.bin`, the objects inside `oleObject*.bin`, and ActiveX binaries.
- VBA source. A VBA project is reported whole, and its modules go unread.
- Encrypted .doc and .ppt files, and Excel's XOR obfuscation. They are rejected.
- Formulas in .xls and .xlsb sheets. A DDE link that a formula uses is found in the workbook's link records.
- Visio .vsd ShapeSheet cells, and anything in Publisher and Project files but VBA projects and objects.
- An Outlook message's body, its RTF body included, and what an attached file holds.
- The piece table of a Word 6.0 or 95 document. Its text is read in the order it is stored between the FIB's bounds, as msodde reads it, so a field that fast save split into pieces stored apart may not read whole.
- Images, fonts and other media, and the inside of an embedded Office package such as a chart's workbook.
- Excel functions other than DDE that reach outside, such as WEBSERVICE, RTD and IMAGE. Data connections in `connections.xml`.
- OpenDocument form submit targets, Word's hyperlink base, and whether a link's visible text names a different site than its target.
- Flat OpenDocument files such as `.fodt`, Flat OPC, and XPS. Inspect rejects them as not a ZIP archive or not a recognized package.
- Bugs in an application's parsers that ordinary content can trigger.

## Development

Development needs Node 22.12 or later. mocha 12 requires it, and the CommonJS test build loads chai 6, an ES module, with `require()`, which Node 22 supports only from 22.12. The pre-commit hook needs Node 22.22.1 or later for lint-staged. The tests also run `python3` and `unzip` to read the ZIP writer's output. The published package has no runtime dependencies and runs on Node 18 or later.

```sh
npm install --ignore-scripts
npm run build
npm test
npm run coverage
npm run lint
npm run format
```

`npm run build` compiles `src` into `dist`. `npm test` runs the lint, builds the package and the tests, then runs mocha, so a lint error fails the tests. `npm run coverage` runs the tests under c8 and prints the line and branch coverage of each source file. `npm run lint` runs `biome check` over the whole tree and changes nothing. `npm run format` rewrites files in Biome's format.

`npx husky` activates the pre-commit hook. A plain `npm install` does the same through the `prepare` script. The hook runs Biome on the staged files, applying its safe fixes, then runs `npm audit`. The commit stops if either one fails.

The fixtures under `test/fixtures` are committed. `python3 -I scripts/make-fixtures.py` rebuilds them, and needs LibreOffice's `soffice`, Python 3 and Info-ZIP's `zip`. It converts flat OpenDocument sources with LibreOffice, which also saves .doc, .xls, .ppt and .rtf versions of some of them. It adds macros, fields, relationships, objects and controls with Python's `zipfile`, and builds the hostile archives byte by byte. Compound files, RTF, Outlook messages and encrypted files come from its own compound file writer, AES and RC4. Two runs write the same bytes: the script sets `SAL_RAND_REPEATABLE` for LibreOffice, and fixes the timestamps and random ids LibreOffice writes.

`test/fixtures/cases` holds one small file for each behavior a run over a corpus of real Office files turned up, written from scratch by `python3 -I scripts/make-fixtures.py cases`, which needs only Python. No corpus file, file name or path is in the repository, macro and OLE payloads are inert markers, and every address uses a reserved name. `manifest.json` gives each file's issue, what office-defuse reports for it, and what the build before the fix it documents reported. A test checks every entry, and that the manifest and the folder list the same files.

## License

MIT
