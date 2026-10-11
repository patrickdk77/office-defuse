# office-defuse

office-defuse reads Word, Excel, PowerPoint, Visio and OpenDocument packages and reports VBA and Excel 4.0 macros, LibreOffice Basic and scripts, DDE, fields and events that run code, templates and objects loaded from outside the file, embedded OLE objects, ActiveX controls and unsafe links. It scores the risk from 0 to 100.

This version inspects and changes nothing.

It has no runtime dependencies and keeps memory bounded. It reads the ZIP and XML containers strictly, and refuses anything two readers could take in different ways.

```
npm install @patrickdk77/office-defuse
```

Node 22 or later.

## Quick start

```js
const { inspectOffice } = require('@patrickdk77/office-defuse');

const report = await inspectOffice(bytes);
report.status;   // 'clean', 'strippable' or 'rejected'
report.risk;     // 'NONE', 'LOW', 'MEDIUM', 'HIGH', 'CRITICAL' or 'UNKNOWN'
report.score;    // 0 to 100, null when the file could not be inspected
report.format;   // 'docx', 'xlsm', 'odt', ... or '' when not recognized
report.mimeType; // the MIME type of that format, or ''
report.findings; // [{ category, detail, description, action, location, data }]
```

ES modules work too: `import { inspectOffice } from '@patrickdk77/office-defuse'`.

`strippable` means at least one finding has the action `strip`, content a defused copy would leave out. `rejected` means a finding has the action `reject`, or the run stopped before it read the whole package.

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

| Family | Formats |
|---|---|
| Word | .docx .docm .dotx .dotm |
| Excel | .xlsx .xlsm .xltx .xltm .xlam .xlsb |
| PowerPoint | .pptx .pptm .potx .potm .ppsx .ppsm .ppam .sldx .sldm |
| Visio | .vsdx .vsdm .vstx .vstm .vssx .vssm |
| OpenDocument | .odt .ott .ods .ots .odp .otp .odg .otg .odf .odm .odb |

`[Content_Types].xml` marks an Office Open XML package. The package relationship of type `officeDocument`, transitional or strict, or Visio's `document`, names the main part, and its content type gives the format. The `mimetype` entry marks an OpenDocument package, or the manifest's entry for `/` when there is no mimetype entry.

A ZIP with both markers is rejected as PACKAGE/AMBIGUOUS_FORMAT, since Office could read one and LibreOffice the other. So is an OOXML package with two main parts, and an ODF package whose mimetype and manifest name different formats. Anything else is PACKAGE/UNRECOGNIZED.

A .xlsb gets the package checks. Its sheets are binary records, and the formulas in them go unread. Its external link parts are binary too, and office-defuse reads their link records for DDE links. A .odb gets every ODF check, and the database under `database/` goes unread.

## What it reports

- VBA projects, Excel 4.0 macro sheets, and OpenDocument `Basic/`, `Scripts/` and `Dialogs/` folders. These go by content type, relationship type or folder name, and a folder counts when it holds a file that is not empty. office-defuse parses nothing inside an OLE file such as `vbaProject.bin`.
- OpenDocument script event listeners that name a script URL or macro. Those that fire as the document opens or is created get a finding of their own. LibreOffice ignores a listener whose URL and macro name are both empty, and so does office-defuse. Visio ShapeSheet cells that call RUNADDON, RUNADDONWARGS, CALLTHIS, RUNMACRO, OPENFILE or DOOLEVERB.
- Links and actions that run macros: `vnd.sun.star.script:` and `macro:` targets, PowerPoint's `ppaction://macro`, Excel drawing shapes with a `macro` attribute, and form controls with an `x:FmlaMacro`. Ribbon callbacks such as `onLoad` and `onAction`. OpenDocument menu, toolbar and key settings under `Configurations2/`. LibreOffice writes an empty `Configurations2/accelerator/current.xml` into the documents it saves, and an empty file counts for nothing.
- PowerPoint and Impress actions that run a program or an object verb.
- DDE: Word `DDE` and `DDEAUTO` fields, Excel `ddeLink` parts and the .xlsb record for them, Excel formulas in `application|topic!item` form, OpenDocument DDE connections and `DDE()` formulas. No amount of white space before a field's name hides it, and a formula is checked whole however long it runs.
- Word fields that run a macro, `MACROBUTTON`, and fields whose type another field builds. That is the QUOTE trick that hides `DDEAUTO`.
- References that load from outside the file on their own: attached and OpenDocument templates, linked OLE objects and packages, images, frames, included files and linked sections, external workbooks, `meta:auto-reload`, and any other external relationship. A target on the machine itself is EXTERNAL/LOCAL_FILE. An empty OpenDocument reference is the document itself and names no file, so it goes unreported. A target counts as local only if its percent-decoded form is local too, because a CVE-2017-0199 sample writes its `script:` scheme as `%73cript:`.
- Hyperlinks, with pdf-defuse's checks of the target: the scheme, file and network paths, data URLs, credentials before the host, IP hosts, punycode and mixed-script hosts, and percent-encoded hosts. A link to a place inside the document goes unreported, and so does an empty link, which is the document itself.
- Embedded OLE objects and packages, OLE files nothing declares as an object, Java applets, plug-ins and ActiveX controls.
- Package structure: parts without a content type, parts no relationship reaches, files the ODF manifest does not list, an irregular ODF layout, and names a package gives twice.
- Document type declarations in XML parts, and references to external entities. See [Document types and entities](#document-types-and-entities).

The OOXML rules match elements by namespace URI and local name, compared ignoring case, so a prefix bound to the Word namespace reads as Word whatever it is called. The OpenDocument rules read `urn:oasis:names:tc:<any>:xmlns:<name>:1.<any>` as version 1.0 of that namespace, which is what LibreOffice does. Every branch of `mc:AlternateContent` gets read.

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

A run that stops rejects the file with a null score and the UNKNOWN level, whatever `actionOverrides` say about the finding that stopped it. The containers, a limit, an unrecognized or ambiguous format, an encrypted OpenDocument package and an unexpected error, PACKAGE/UNREADABLE, all stop a run. The findings made before the stop stay in the report.

## Findings

Every finding has a `category` and a `detail`, plain strings such as `MACRO` and `VBA_PROJECT`, and a `description` that is safe to show the uploader. A description that quotes a value from the file cuts it to 60 characters and shows control and format characters as `?`. `data` keeps a value from the file up to 200 characters, or 2,000 for a target, and `location` names the part, the relationship or the folder.

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

Three combinations add more. Code here means macros or DDE. Code or a run-program action together with a template, object or include from outside adds 20. Code with an embedded object or ActiveX control adds 15. Code with leading, hidden or trailing archive data adds 10. The total is capped at 100.

The bands are NONE at 0, LOW from 1, MEDIUM from 25, HIGH from 50 and CRITICAL from 80. `scoreWeights` and `scoreBands` change them. A finding of a kind office-defuse does not make, such as one from another package, scores at its `weight`, or 0 without one.

## Options

| Option | Default | Purpose |
|---|---|---|
| `limits.fileSize` | none | Bytes in the upload, checked before any read |
| `limits.entries` | 10,000 | Entries in the central directory |
| `limits.decompressedBytes` | 512 MiB | The declared size of any one entry |
| `limits.totalDecompressedBytes` | 2 GiB | The declared sizes of all entries together |
| `limits.xmlDepth` | 256 | Element nesting in one XML part |
| `limits.xmlTokenSize` | 16 Mi | Characters in one tag, comment, processing instruction or CDATA section. A document type declaration counts as one. Text goes through in pieces and has no limit |
| `limits.xmlEntityDepth` | 16 | Entity references nested inside entity text |
| `limits.xmlEntityExpansions` | 100,000 | Entity references expanded, and attribute defaults added, in the whole run |
| `limits.xmlEntitySize` | 16 Mi | Characters of entity text and attribute defaults expanded in the whole run |
| `limits.timeMs` | none | The whole run |
| `actionOverrides` | none | Change the action for a category, or for a category and detail |
| `scoreWeights`, `scoreBands` | built in | Change the scoring |

The eight limits with a default bound the memory or work any input can cost, so they apply unless you set them. `Infinity` turns one off. A limit that is reached stops the run with a LIMIT finding.

Besides one XML token, the stack of open elements and the declarations of the part's internal subset, office-defuse holds the central directory's table of entries, which parts each part's relationships reach, and, while it reads a relationships part, the Ids that part has given so far, about 100 bytes each, so it can find one given twice. Content types, manifest paths and object references that name no entry of the archive are not kept. Everything else streams.

## The defuse family

office-defuse is one of a family of packages that each handle one kind of file, beside [pdf-defuse](https://github.com/patrickdk77/pdf-defuse). It declares the family's shared types with the same shapes as pdf-defuse: `DefuseFinding`, `ByteSource`, `ByteSink`, `PluginContext`, `ContainedFile`, `ContainedFileResult` and `ContainedFilePlugin`. It does not depend on pdf-defuse, and TypeScript matches the types by shape.

A finding's `category` and `detail` are plain strings, and `weight`, when set, says how much it adds to a score. The time limit's error carries `code: 'DEFUSE_LIMIT'` and `limit: 'time'`, the code every package of the family passes up the tree of files. `inspectOffice` turns it into LIMIT/TIME.

## What it does not check

- The contents of OLE files, such as VBA source, the objects inside `oleObject*.bin`, and ActiveX binaries.
- Images, fonts and other media in a package, and the inside of an embedded Office package such as a chart's workbook.
- Formulas in .xlsb sheets, and Excel functions other than DDE that reach outside, such as WEBSERVICE, RTD and IMAGE. Data connections in `connections.xml`.
- OpenDocument form submit targets, Word's hyperlink base, and whether a link's visible text names a different site than its target.
- Flat OpenDocument files such as `.fodt`, Flat OPC, and XPS. Inspect rejects them as not a ZIP archive or not a recognized package.
- Bugs in an application's parsers that ordinary content can trigger.

## Development

Development needs Node 22.12 or later. mocha 12 requires it, and the CommonJS test build loads chai 6, an ES module, with `require()`, which Node 22 supports only from 22.12. The pre-commit hook needs Node 22.22.1 or later for lint-staged. The tests also run `python3` and `unzip` to read the ZIP writer's output. The published package has no runtime dependencies and runs on any Node 22.

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

The fixtures under `test/fixtures` are committed. `python3 -I scripts/make-fixtures.py` rebuilds them, and needs LibreOffice's `soffice`, Python 3 and Info-ZIP's `zip`. It converts flat OpenDocument sources with LibreOffice, adds macros, fields, relationships, objects and controls with Python's `zipfile`, and builds the hostile archives byte by byte.

`test/fixtures/cases` holds one small package for each behavior a run over a corpus of real Office files turned up, written from scratch by `python3 -I scripts/make-fixtures.py cases`, which needs only Python. No corpus file, file name or path is in the repository, macro and OLE payloads are inert markers, and every address uses a reserved name. `manifest.json` gives each file's issue, what office-defuse reports for it, and what the build before the fix it documents reported. A test checks every entry, and that the manifest and the folder list the same files.

## License

MIT
