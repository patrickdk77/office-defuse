# office-defuse

office-defuse reads Word, Excel, PowerPoint and Visio files, in their ZIP formats and in the compound file formats before them, OpenDocument packages, Outlook messages, Publisher and Project files, and RTF. It reports VBA, WordBasic and Excel 4.0 macros, LibreOffice Basic and scripts, DDE, fields and events that run code, templates and objects loaded from outside the file, embedded OLE objects, Equation Editor objects, attachments, ActiveX controls and unsafe links. It scores the risk from 0 to 100.

It also writes defused copies. A copy leaves out what each strip finding names, and it reaches you only after it passes an inspection of its own.

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

`strippable` means at least one finding has the action `strip`, content a defused copy would leave out, or a file plugin cleaned a file inside. `rejected` means a finding has the action `reject`, or the run stopped before it read the whole file.

```js
const { disarmOffice } = require('@patrickdk77/office-defuse');

const result = await disarmOffice(bytes);
result.status;  // 'clean', 'defused' or 'rejected'
result.before;  // the inspection of the upload
result.after;   // the inspection of the output, absent when rejected
result.removed; // the findings whose content the output leaves out
result.bytes;   // the output, absent when rejected
```

`clean` means the upload had nothing to remove and the output is a rewrite of it. See [Disarm](#disarm).

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

`disarmOfficeSource(source, sink, options)` writes the output to a sink with `write(chunk)`, `close()` and an optional `abort(error)`. `bufferSink`, `fileSink` and `writableSink`, which wraps a Node stream, are included. The run copies the upload once, into memory up to `memoryThreshold` and into a temporary file past it, and reads only that copy, so a source that changes during the run cannot change the output. Nothing reaches the sink until the output has passed its check. A copy to the sink that fails calls `abort()` instead of `close()`, and `fileSink` then deletes the partial file.

```js
const { disarmOfficeSource, fileSource, fileSink } = require('@patrickdk77/office-defuse');

const result = await disarmOfficeSource(fileSource('upload.docx'), fileSink('clean.docx'));
```

## Formats

The format comes from the content, never from the file name.

| Family | ZIP formats | Compound file and RTF formats | Flat XML formats |
|---|---|---|---|
| Word | .docx .docm .dotx .dotm | .doc .dot from Word 6.0 to 2003, and .rtf | Flat OPC .xml |
| Excel | .xlsx .xlsm .xltx .xltm .xlam .xlsb | .xls .xlt .xla from Excel 5.0 to 2003 | Flat OPC .xml |
| PowerPoint | .pptx .pptm .potx .potm .ppsx .ppsm .ppam .sldx .sldm | .ppt from PowerPoint 97 to 2003 | Flat OPC .xml |
| Visio | .vsdx .vsdm .vstx .vstm .vssx .vssm | .vsd from Visio 2000 to 2010 | |
| XPS | .xps .oxps | | |
| OpenDocument | .odt .ott .ods .ots .odp .otp .odg .otg .odf .odm .odb | | .fodt .fods .fodp .fodg |
| Outlook | | .msg .oft | |
| Publisher | | .pub | |
| Project | | .mpp | |

`[Content_Types].xml` marks an Office Open XML package. The package relationship of type `officeDocument`, transitional or strict, or Visio's `document`, names the main part, and its content type gives the format. An XPS package names its FixedDocumentSequence with a `fixedrepresentation` relationship. XPS and OpenXPS share their content types, so the relationship's namespace tells them apart. The `mimetype` entry marks an OpenDocument package, or the manifest's entry for `/` when there is no mimetype entry.

A ZIP with both markers is rejected as PACKAGE/AMBIGUOUS_FORMAT, since Office could read one and LibreOffice the other. So is an OOXML package with two main parts, and an ODF package whose mimetype and manifest name different formats. Anything else is PACKAGE/UNRECOGNIZED.

A .xlsb gets the package checks. Its sheets are binary records, and the formulas in them go unread. Its external link parts are binary too, and office-defuse reads their link records for DDE links. A .odb gets every ODF check, and the database under `database/` goes unread.

A file that starts with the compound file signature is a compound file, MS-CFB. The streams and storages at its root name the format: `WordDocument` for Word, `Workbook` or `Book` for Excel, `PowerPoint Document`, `VisioDocument`, `__properties_version1.0` for an Outlook message, a `Quill` storage or Publisher's class id beside a `Contents` stream for Publisher, and a `Props` stream beside a storage named by a number for Project. Word's FIB says whether a document is a template, and Excel's records whether a workbook is a template or an add-in. Neither needs the password of an encrypted file. An Outlook template has a class id of its own. A PowerPoint 97 to 2003 template or show reports its format as .ppt, and a Visio stencil or template as .vsd.

A compound file with two of these is PACKAGE/AMBIGUOUS_FORMAT. One with none is PACKAGE/UNRECOGNIZED, and so is a `WordDocument` stream that holds no Word 6.0 to 2003 FIB. A PowerPoint 4.0 file, a bare OLE object and a `vbaProject.bin` on its own all land there.

An encrypted .docx, .xlsx or .pptx is a compound file holding `EncryptionInfo` and `EncryptedPackage`. See [Encrypted files](#encrypted-files).

A file that starts as XML, after any byte order mark and white space, is a flat file. A root `office:document` makes it flat OpenDocument, whose `office:mimetype` gives the format. A root `pkg:package` makes it Flat OPC, the one-file form Word, Excel and PowerPoint save as XML, and its format is `xml` whatever its main part is. Any other root is PACKAGE/UNRECOGNIZED, and so is a flat OpenDocument media type outside the table. A file that is XML and a ZIP archive too is PACKAGE/AMBIGUOUS_FORMAT.

A file that starts with `{\rt` is RTF, as Word takes it. A compound file or RTF file in which the ZIP reader also finds an archive is PACKAGE/AMBIGUOUS_FORMAT. A file with the compound file signature whose compound file does not read, but whose ZIP archive does, is that ZIP archive with LEADING_BYTES.

## Detection

`detectOffice(bytes, options)` and `detectOfficeSource(source, options)` tell a router whether office-defuse reads a file. They return `{ mimeType, extension, offset }`, or undefined. `extension` and `mimeType` are the `format` and `mimeType` an inspection reports. The inspection and detection name a file with the same functions. Detection opens the container and reads only what names the format, without the rules, the plugins or the rewrite:

- the streams at a compound file's root, and for Word and Excel the FIB or the workbook's globals, which say whether it is a template or an add-in
- a package's `[Content_Types].xml` and package relationships, or its `mimetype` entry and manifest
- a flat file's root element, and for Flat OPC each part's name and content type and the package relationships
- the first bytes of RTF

Detection does not claim a file whose format office-defuse cannot name. That covers a ZIP archive with no Office marker and a package of another kind, such as a NuGet package or a StarOffice file. It covers a compound file with no Office document at its root, such as an installer or a bare OLE object, and a file that is a ZIP archive and a compound file at once. An encrypted OOXML package names its format only once it is decrypted, and detection has no password. When VelvetSweatshop opens the package, detection names the format inside. Otherwise it claims the package as a compound file, application/x-ole-storage with an empty extension, since the file is office-defuse's whatever the password turns out to be. The inspection then decrypts it with the run's password, or rejects it as encrypted. Detection claims a file it can name even when the file is damaged further on, and the inspection then rejects it.

`offset` is where the ZIP data starts in an archive with other data in front, as the ZIP reader finds it, and 0 for anything else. The name and declared type play no part, so `heuristic` is never set. `detectOfficeSource` leaves the source open and throws with the `DEFUSE_LIMIT` code once `options.deadline` passes. An I/O error from the source comes back unchanged. Naming a Flat OPC package packs its parts into a ZIP archive, which spills into `options.tempDir` past `options.memoryThreshold`. `officePlugin` accepts exactly what detection claims.

## What it reports

In a ZIP package:

- VBA projects, Excel 4.0 macro sheets, and OpenDocument `Basic/`, `Scripts/` and `Dialogs/` folders. These go by content type, relationship type or folder name, and a folder counts when it holds a file that is not empty. office-defuse parses no OLE part of a package, such as `vbaProject.bin`, unless `scriptPlugins` is set. Then it reads the VBA project, as [Script plugins](#script-plugins) says.
- OpenDocument script event listeners that name a script URL or macro. Those that fire as the document opens or is created get a finding of their own. LibreOffice ignores a listener whose URL and macro name are both empty, and so does office-defuse. Visio ShapeSheet cells that call RUNADDON, RUNADDONWARGS, CALLTHIS, RUNMACRO, OPENFILE or DOOLEVERB.
- Links and actions that run macros: `vnd.sun.star.script:` and `macro:` targets, PowerPoint's `ppaction://macro`, Excel drawing shapes with a `macro` attribute, and form controls with an `x:FmlaMacro`. Ribbon callbacks such as `onLoad` and `onAction`. OpenDocument menu, toolbar and key settings under `Configurations2/`. LibreOffice writes an empty `Configurations2/accelerator/current.xml` into the documents it saves, and an empty file counts for nothing.
- Excel defined names that run a macro sheet: Auto_Open and Auto_Activate as MACRO/EVENT_ON_OPEN, Auto_Close and Auto_Deactivate as MACRO/EVENT. A name counts when it starts with one of these, as `Auto_Open2` does.
- PowerPoint and Impress actions that run a program or an object verb.
- DDE: Word `DDE` and `DDEAUTO` fields, Excel `ddeLink` parts and the .xlsb record for them, Excel formulas in `application|topic!item` form, OpenDocument DDE connections and `DDE()` formulas. No amount of white space before a field's name hides it, and a formula is checked whole however long it runs.
- Word fields that run a macro, `MACROBUTTON`, and fields whose type another field builds. That is the QUOTE trick that hides `DDEAUTO`.
- Formula functions that reach outside the file when Excel or LibreOffice calculates them: WEBSERVICE and IMAGE fetch a URL, as EXTERNAL/FORMULA, and RTD, CALL, REGISTER, REGISTER.ID and EXEC run code outside the file, as MACRO/FORMULA. A prefix such as `_xlfn.` or `COM.MICROSOFT.` names the same function. A name in quotes is text, and a macro sheet's formulas are its macros, which its own finding covers. The rewrite removes the formula and the cell keeps its cached value.
- The sources an Excel data connection in `connections.xml` loads from, each as EXTERNAL/DATA_CONNECTION: an ODC file or offline cube, a connection string and its command, a web query's URL, or a text file. The rewrite empties each source and keeps the connection, which query tables and pivot caches name.
- Where an OpenDocument form sends its fields, as ACTION/SUBMIT_FORM, which the rewrite removes, as pdf-defuse removes a form submit. Word's hyperlink base in `docProps/app.xml` gets the link checks.
- References that load from outside the file on their own: attached and OpenDocument templates, linked OLE objects and packages, images, frames, included files and linked sections, external workbooks, `meta:auto-reload`, and any other external relationship. A target on the machine itself is EXTERNAL/LOCAL_FILE. An empty OpenDocument reference is the document itself and names no file, so it goes unreported. A target counts as local only if its percent-decoded form is local too, because a CVE-2017-0199 sample writes its `script:` scheme as `%73cript:`.
- Hyperlinks, with pdf-defuse's checks of the target: the scheme, file and network paths, data URLs, credentials before the host, IP hosts, punycode and mixed-script hosts, and percent-encoded hosts. A link to a place inside the document goes unreported, and so does an empty link, which is the document itself.
- A safe link whose label is an address of another site than the one it opens is LINK/TEXT_MISMATCH. A label counts as an address when it is a web URL, a mail address, or a host with or without a path that starts with `www.` or ends in a known top-level domain. Text that only mentions an address names no site, and neither do Excel function names such as `BETA.DIST` and file names such as `moon.htm`, which look like hosts. Sites compare as in pdf-defuse: a subdomain or parent of the label's site matches, and `evil.co.uk` is not under `mybank.co.uk`. The labels are a Word hyperlink's text and tooltip, a HYPERLINK field's result and `\o` tooltip, a DrawingML run's text and a link's tooltip, an Excel hyperlink's display name and tooltip, an OpenDocument link's text and title, and the characters of an XPS Glyphs element. The rewrite removes the link and keeps its text.
- Embedded OLE objects and packages, OLE files nothing declares as an object, Java applets, plug-ins and ActiveX controls.
- Package structure: parts without a content type, parts no relationship reaches, files the ODF manifest does not list, an irregular ODF layout, and names a package gives twice.
- Document type declarations in XML parts, and references to external entities. See [Document types and entities](#document-types-and-entities).

- In XPS, the markup names the parts it uses: a document's, page's or resource dictionary's `Source`, a font's `FontUri`, an image brush's `ImageSource` and a `ContextColor` profile. Each counts as reaching its part. One outside the package is EXTERNAL, with IMAGE for an image, INCLUDE for a `Source` and OTHER for the rest, and the element that names it goes. A `FixedPage.NavigateUri` or `OutlineTarget` with a scheme or a network path gets the link checks, and a relative one moves inside the document. Pictures go to the file plugins, and obfuscated fonts, which are part of the format, do not.

The OOXML rules match elements by namespace URI and local name, compared ignoring case, so a prefix bound to the Word namespace reads as Word whatever it is called. The OpenDocument rules read `urn:oasis:names:tc:<any>:xmlns:<name>:1.<any>` as version 1.0 of that namespace, which is what LibreOffice does. Every branch of `mc:AlternateContent` gets read.

A flat OpenDocument file gets the OpenDocument rules over its one part, which findings name `office:document`. It has no entries, so every relative reference names a file beside the document and loads from outside it, as EXTERNAL/LOCAL_FILE. A chart's `xlink:href=".."` names the document that holds the chart, and a `vnd.sun.star.Package:` URL the document's own storage, so neither counts. The pictures and OLE objects it holds as base64 in `office:binary-data` go to the file plugins, and an OLE object no plugin keeps is EMBEDDED/OLE_OBJECT.

A Flat OPC package is unpacked into a package of its parts, each with its `pkg:contentType`, and read with the OOXML rules. A part's XML keeps the namespaces it inherits from `pkg:package`. Two parts with one name are CONTAINER/DUPLICATE_NAME, and a part with no name is PACKAGE/UNRECOGNIZED.

In a compound file, the first two rules run over every storage of the tree, and the others over the streams of the format:

- VBA projects: a `VBA` storage with a `dir` stream, or a `_VBA_PROJECT_CUR` or `Macros` storage with something in it.
- Embedded objects, by the streams an object's storage holds, under `ObjectPool` or anywhere else: OLE objects with their class id, ProgID and user type, the file a Packager object holds, Equation Editor 3.0 objects by class id or ProgID, ActiveX controls, and linked objects with the target their moniker names. An object's storage is part of the tree, so a macro or an object inside an object is found.
- Word: WordBasic macros in Word 6.0 and 95 files. The fields of the text, with the rules a .docx gets, and the attached template. A Word 6.0 or 95 file saved fast is read in the order its piece table gives, and any other one from fcMin to fcMac.
- Excel: Excel 4.0 macro sheets, by sheet type or by the sheet's own BOF record, and the Auto_ names above, built in or defined. The functions above in a worksheet's cell formulas, by their index in Excel's function table or by the name a newer function is called by, such as `_xlfn.WEBSERVICE`. External workbooks, with the paths Excel encodes written out, DDE and OLE links, hyperlinks, shapes and buttons that run a macro, and the `Ctls` stream that holds ActiveX controls.
- PowerPoint: actions that run a macro, a program or an object verb, hyperlinks, and linked sounds and movies. PowerPoint keeps the storages of objects, controls and the VBA project in its document stream, often compressed. office-defuse inflates them and gives them the rules above.
- Outlook: each attachment, as EMBEDDED/ATTACHMENT with its name, MIME type and size, and each attachment by reference with its path. An attached message's own attachments are read, 16 messages deep. The RTF body, decompressed as MS-OXRTFCP says, gets the RTF rules, and its findings name `RTF body` as their attachment. A body that does not decompress, or whose CRC does not match, is PACKAGE/UNREADABLE_BODY. Each embedded message's bodies get the same reading.
- Visio 2000 to 2010: ShapeSheet cells whose formulas call RUNADDON, RUNADDONWARGS, CALLTHIS, RUNMACRO, OPENFILE or DOOLEVERB, as in a .vsdx. A call in EventDblClick, EventXFMod, EventDrop or EventMultiDrop is MACRO/EVENT, and one in any other cell MACRO/LINK, with the section, the cell and the formula. Microsoft has not published the format, so office-defuse reads it as libvisio, Apache POI and OLEToy describe it, and finds every call to these functions in their test files. Visio stores a formula as tokens. The finding writes the formula out when it holds only constants, operators and the call, and gives the function's name alone otherwise. office-defuse then searches each stream for the token of such a call outside the formulas it read. That covers chunk types whose cells it cannot place, and the finding names the cell `unknown`. Decompressed streams count toward `limits.decompressedBytes` and `limits.totalDecompressedBytes`. A pointer, chunk, cell block or formula that does not read stops the run as PACKAGE/UNREADABLE. The VisioDocument stream of Visio 5.0 and earlier goes unread.
- Publisher and Project: the VBA projects and objects. Nothing in their own streams is read.

In RTF:

- Objects in `\objdata`. office-defuse decodes the OLE 1.0 object, the file a Packager object holds, and the storage of an OLE 2 object, which gets the compound file rules. An object linked with `\objlink` or `\objautlink`, or of the class OLE2Link that CVE-2017-0199 uses, is EXTERNAL/OBJECT. `\objocx` marks an ActiveX control. `\objupdate` is ACTION/OBJECT_UPDATE, since it starts the object's program as the document opens.
- Fields, with the rules a .docx gets, the `\*\template` destination, and shape hyperlinks.
- Text after the group that closes the document, as TRAILING_BYTES. White space and NULs there are what writers leave, and count for nothing.

### Encrypted files

office-defuse decrypts these, all from MS-OFFCRYPTO:

- An OOXML package encrypted with Standard or Agile encryption, Agile in CBC or CFB chaining.
- An .xls workbook encrypted with RC4 or RC4 CryptoAPI, or obfuscated with XOR, Excel 5.0 and 95's only kind.
- A Word 97 to 2003 document encrypted with RC4 or RC4 CryptoAPI, or obfuscated with XOR. Its WordDocument, table and Data streams are decrypted where they are read.
- A PowerPoint 97 to 2003 presentation encrypted with RC4 CryptoAPI, one persist object at a time, each with the key of its persist id.

`password` takes one password or a list, since the files in one upload can have different passwords. office-defuse tries each one in order, each once, and then the password the application tries on its own: VelvetSweatshop for Excel and OOXML, and `/01Hannes Ruescher/01` for PowerPoint, so such a file opens without asking. Word has none. A password meant for another file never keeps the default from opening this one. The decrypted file gets every check of its format. It is read in place. Agile encryption decrypts in 4,096-byte segments and Standard in 16-byte blocks, so only the parts office-defuse reads get decrypted. A PowerPoint stream is decrypted whole into memory, up to `limits.decompressedBytes`.

A binary file stays encrypted in the output, since the rewrite copies its streams as they are. The output then still holds ENCRYPTED/DEFAULT_PASSWORD or USER_PASSWORD, which strip, so disarm rejects it unless an override makes them `info`. An encrypted OOXML package comes out decrypted.

| Finding | Action | Weight | When |
|---|---|---|---|
| ENCRYPTED/DEFAULT_PASSWORD | `strip` | 10 | An application's default password opened the file. Anyone can open it, so the encryption hides its content from a scanner and from no one else |
| ENCRYPTED/USER_PASSWORD | `strip` | 0 | The password the caller gave opened the file |
| ENCRYPTED/RC4_40, RC4_128, RC4_OTHER, AES_128, AES_192, AES_256, XOR_OBFUSCATION | `info` | 0 | The cipher and key size |
| ENCRYPTED/PASSWORD_REQUIRED | `reject` | 0 | No password at hand opens the file, or an Agile spin count is above 10,000,000. Every encrypted Word 6.0 or 95 document, which office-defuse does not decrypt |
| ENCRYPTED/UNSUPPORTED_ALGORITHM | `reject` | 0 | A cipher, chaining or hash outside the list below |

AES-128, AES-192 and AES-256 in CBC or ECB mode, SHA-1, SHA-256, SHA-384, SHA-512 and MD5 go to Node's crypto module, and no other algorithm name from a file reaches it. Agile's CFB chaining, CFB with an 8-bit window, is built on AES-ECB. RC4 and XOR obfuscation are office-defuse's own code, since OpenSSL 3 keeps RC4 in its legacy provider.

Pass passwords in code, or to the command-line tool in a file or `OFFICE_DEFUSE_PASSWORD`. Never pass one as an argument, where other users can read it.

## Disarm

`disarmOffice` and `disarmOfficeSource` inspect the file, write a copy without the content each strip finding names, and inspect the copy. If that check reports any finding that is not `info`, the copy is thrown away. The upload is rejected with PROCESSING/VERIFICATION_FAILED, whose `data.reasons` names what the check found, and it keeps its score. A copy whose format differs from the upload's fails the same way. Every file that is not rejected is rewritten, clean ones included, so the output holds only what the run read.

The check reads the output with the caller's overrides and limits, without `fileSize`, `timeMs` and `nestingDepth`, and within the deadline the upload's run has. Disarm reads the file twice, once to inspect it and once to check the output, so it takes about twice as long as an inspection.

### ZIP packages

- The archive is new. Data before the first entry, between entries and after the end goes, and so do extra fields and data descriptors. An entry nothing changes keeps its compressed bytes and its timestamp.
- A part goes with its own relationships, every relationship that targets it, its override in `[Content_Types].xml`, and the parts that only it reached, so a VBA project takes its signature and `vbaData.xml` with it.
- A relationship a strip finding names goes, such as an external template, object, image or frame, or a hyperlink to an unsafe target.
- An element that names a relationship that went goes too, so no element points at nothing. A `w:hyperlink` loses its tags and keeps its text. A SpreadsheetML list that loses its last item goes. An external link part whose link went goes, and the workbook's reference to it with it.
- A part loses its root element only by going whole. Edits happen in place, and everything the rewrite does not touch keeps its characters, in the part's own encoding, UTF-8 or UTF-16, behind its own byte order mark.
- A Word field a strip finding names loses its begin, separator and end marks and its code, nested fields included, and keeps its result. A `w:fldSimple` keeps the runs inside it.
- An Excel 4.0 macro sheet becomes an empty worksheet of the same name, with the worksheet content type and relationship type, so the indexes of the sheets after it stay right. A defined name that runs a macro goes. So does a formula that calls DDE, and its cell keeps its cached value.
- A drawing shape loses its `macro` attribute, a ribbon control its callbacks, and a PowerPoint shape an action that runs a program, a macro or an object verb. A Visio cell that calls code goes. When script plugins keep the document's VBA project, a shape, form control, callback, action or MACROBUTTON field that runs one of its macros by name stays.
- The main part keeps its content type, so a `.docm` without its VBA project is still a `.docm`.
- In an OpenDocument package, `mimetype` comes first, stored, without an extra field, and holds the media type without white space. A package without one gets one, and a package without a manifest gets a manifest listing its entries. `Basic/`, `Scripts/`, `Dialogs/` and `Configurations2/` go whole. An entry goes with its manifest entry and with every element whose `xlink:href` names it. A link `text:a` or `draw:a` keeps what it holds. A form control loses its link, and a cell loses a formula that calls DDE. An embedded OLE object goes and the preview image beside it stays.
- A document type declaration with an internal subset goes when XML/INTERNAL_SUBSET strips it. Each reference to an internal entity is written out as its text, markup included, and each start tag whose attributes the declarations changed or added to is written out whole, namespace declarations included. The part then reads as the inspection read it. References to external entities go when XML/EXTERNAL_ENTITY strips them. If that finding is overridden to `info` while the subset would go, the references still need their declarations, so the declaration stays and the check of the output reports it.
- An element that came from entity text has the offsets of the reference, so the rewrite writes that one reference out, without the element.
- An encrypted package is written decrypted. office-defuse does not encrypt.
- No PDF header starts in the first 1024 bytes of the copy, where PDF readers look for one. Entries that went, and parts rewritten to another size, move what follows forward, so a stored PDF, an entry name or a local header from deep in the upload can land there. A copied entry that would put a header there is deflated again. Any other entry that would waits until later entries have filled those bytes, up to 16 entries. What still lands there fails the output check, and the upload is rejected.

### Flat files

A flat OpenDocument file is rewritten in place, as a package's XML part is. An element goes or loses an attribute where a strip finding names it, and base64 that a plugin rewrote replaces the file's own.

A Flat OPC package is rewritten as the package of its parts is, then written back as one XML file in UTF-8. It starts with an XML declaration and the prolog's `mso-application` processing instructions, which tell Windows which application opens the file. Each part keeps its name, its content type and its other attributes, such as `pkg:compression`. XML parts go in `pkg:xmlData` and the rest as base64 in `pkg:binaryData`, in lines of 76 characters. Anything else in the prolog, such as a document type declaration, is left out.

### Compound files

The rewrite writes a new compound file, version 3 with 512-byte sectors, without the streams and storages a strip finding names: VBA project storages, the storages of embedded objects, linked objects and ActiveX controls, Excel's `Ctls` stream, and Outlook attachments. A VBA project that script plugins keep is written without its compiled code. A message's property stream counts its attachments, and the count drops by the ones that went.

The FAT comes right after the header, so no stream lands in the first 1024 bytes, where PDF readers look for a header. When a stream ends with a ZIP end record, as an embedded package does, free sectors of zeros follow the last stream until the record is more than 65,557 bytes from the end of the file, past where ZIP readers look.

An Outlook message's RTF body with a strip finding is rewritten the way an RTF file is, and stored back uncompressed, which MS-OXRTFCP allows. A body that does not decompress, or one with a `reject` finding inside, goes, and the message keeps its plain text body. The property stream entry of a body that went goes too, and that of a rewritten body takes its new size. Every other stream is copied unchanged. Sectors no stream holds, the rest of each stream's last sector and data past the last sector are not written.

Content inside a stream stays as it was. That covers Word's fields, template and WordBasic macros, Excel's records for macro sheets, defined names, links, hyperlinks and buttons, every record of a PowerPoint presentation along with the objects stored in its records, the ShapeSheet cells of a Visio drawing, and the RC4 encryption of a workbook. A finding there is still in the output, so the output fails its check and the upload is rejected.

### RTF

The rewrite copies the file without the groups of the destinations a strip finding names. A field loses its `\fldinst` group and keeps its result. The `\*\template` destination goes, and so does a shape hyperlink's `\hlsrc`. An object's group goes, and the content of its `\result` group takes its place as a plain group, so the picture Word shows stays. When ACTION/OBJECT_UPDATE strips and nothing else does, `\objupdate` goes on its own. Text after the group that closes the document goes.

## File plugins

`filePlugins` takes contained-file plugins with pdf-defuse's shapes. office-defuse shows them each file it finds inside the file:

- In an OOXML package, its OLE parts, which are VBA projects, embedded objects, ActiveX binaries and any part whose bytes are an OLE file, its embedded packages, and every other part that is a whole file in a format of its own, such as the images under `media/` and a thumbnail. The binary parts of the format itself go to no plugin: the records of a binary workbook, printer settings, obfuscated fonts and VBA project signatures.
- In an OpenDocument package, each entry outside `META-INF/` and the code and settings folders that is not read as XML, such as an OLE object, a picture or the thumbnail.
- In an Outlook message, its RTF body, decompressed, and its HTML body. A plugin's copy of the RTF body is stored uncompressed.
- In a compound file the rewrite writes, each embedded object. A Packager object shows the file it holds, and any other object its storage, written out as a compound file of its own. Outlook attachments held in a stream go too, whatever their type. In a PowerPoint presentation, the storages of its objects go as the compound files they are.
- In RTF, an embedded object's native data, or the file a Packager object holds, and each picture that is a whole file, PNG, JPEG or EMF.
- In a flat OpenDocument file, each picture and OLE object held as base64 in `office:binary-data`. A Flat OPC package's parts go as a package's parts do.

A file comes as a `ContainedFile`. `name` is the part or entry name, a Packager file's name, or an object's type. `declaredType` is a part's content type or a picture's type. `sniffedType` comes from the first bytes, with the MIME types pdf-defuse gives and `application/rtf`. `location` says where the file sits, `depth` is its depth in the upload, and `source` reads its bytes. The first plugin whose `accepts` returns true decides.

| The plugin returns | The file | Finding |
|---|---|---|
| `passed` | stays, as the plugin wrote it when it wrote anything | EMBEDDED/PLUGIN_PASSED, `info` |
| `scrubbed`, with bytes written | is replaced by those bytes | EMBEDDED/PLUGIN_SCRUBBED, `info` |
| `removed` | goes, with every reference to it | EMBEDDED/PLUGIN_REMOVED, `strip` |
| `scrubbed` with nothing written, or an error | goes, with every reference to it | EMBEDDED/PLUGIN_FAILED, `strip` |

A file goes whatever the overrides say about these findings. A file no plugin accepts gets the rule it gets without plugins, such as EMBEDDED/OLE_OBJECT, and a picture, which no rule names, stays. When a plugin keeps a file, the findings that name the file for being there, such as EMBEDDED/PACKAGE or MACRO/VBA_PROJECT for `vbaProject.bin`, turn to `info` and leave the score, as pdf-defuse leaves out an attached file a plugin kept. Other rules still decide for themselves, so an Equation Editor object a plugin passes goes anyway, since EMBEDDED/EQUATION strips it. A plugin's output for an object's storage must be a compound file, whose storages and streams then take the storage's place. Where the rewrite cannot change the container, as for an object stored in a PowerPoint record, the output keeps the file as it arrived, so a plugin's `scrubbed` or `removed` there fails the output check.

The findings a plugin returns go into the report with their `location` after the file's own and with `attachment` naming the file. A `reject` inside the file becomes `strip`, since the file goes and the file holding it stays. Each scores at its `weight`. The upload's score leaves out PLUGIN_PASSED and PLUGIN_SCRUBBED, and the output's counts them. The check of the output does not ask the plugins again. It matches each file by the MD5 of its bytes against what the plugins kept, and reports the `outputFindings` the plugin gave for it.

The plugin context has `depth`, the file's depth in the upload, and `deadline`, which every file of the upload shares. An error with `code: 'DEFUSE_LIMIT'` or `code: 'DEFUSE_IO'` from a plugin stops the whole upload. At depth 0, a limit becomes LIMIT/TIME or LIMIT/NESTING_DEPTH, and an I/O error is thrown as its cause.

`officePlugin(options)` defuses Office and OpenDocument files inside other files. It accepts the files [`detectOffice`](#detection) claims, by their content alone, an encrypted OOXML package included, which it decrypts with the run's password. It runs disarm at the context's depth with the context's deadline. Without `options` it takes the options of the office-defuse run that found the file, or none when another package found it. A clean file comes back `passed`, a defused one `scrubbed`, both with the rewritten bytes, and a rejected one `removed`. The findings it returns carry their `weight`.

```js
const { disarmOffice, officePlugin } = require('@patrickdk77/office-defuse');

await disarmOffice(bytes, { filePlugins: [officePlugin()] });
```

## Script plugins

`scriptPlugins` takes script plugins that decide about VBA. office-defuse reads a VBA project's dir stream and modules as MS-OVBA describes them, and shows each module to the plugins as a `ContainedScript`:

| Field | What it holds |
|---|---|
| `text` | The module's source, decoded with the project's code page |
| `trigger` | `vba-document`, `vba-standard`, `vba-class` or `vba-form`, as the PROJECT stream lists the module |
| `location` | Where the module sits, such as `xl/vbaProject.bin > VBA/Module1` |
| `fieldName` | The module's name |
| `language`, `form` | `vba` and `office` |
| `references` | On the first module only, the project's references, each `{ name, kind, libid }` |

The first plugin whose `accepts` returns true decides a module. A module it passes gets MACRO/PLUGIN_PASSED, `info`. Any other answer, `scrubbed` included, is MACRO/PLUGIN_REMOVED, and an error is MACRO/PLUGIN_FAILED, both `strip`, since office-defuse writes no module a plugin rewrote. The plugin's findings go into the report as they are.

The project stays when every module passed. Otherwise it goes, as it does without plugins, and so does a project with a module no plugin took. A project that stays is written without its compiled code, the way Mandiant's OfficePurge writes one. Each module stream holds only its source, compressed again, the dir stream says each source starts at 0, `_VBA_PROJECT` names no VBA version, and the `__SRP_` caches go. Office then compiles the source the plugins read, so p-code that does something other than its source, which VBA stomping relies on, never runs. MACRO/COMPILED_CODE, `strip`, reports the purge, and MACRO/VBA_PROJECT turns to `info`.

A user form stays with its project when its designer storage holds Forms 2.0 controls only. The designer must be the UserForm designer, the storages under it must be Forms 2.0 forms, frames and multipages, and every form's class table must be empty, since a class table can name any ActiveX control. A form that fails gets CONTROL/ACTIVEX at its designer, and while that finding strips, the project goes.

When the document's own project stays, so does what runs its macros by name: an Excel shape's `macro`, a form control's `x:FmlaMacro`, a ribbon callback, PowerPoint's `ppaction://macro` and Word's MACROBUTTON field. Their findings turn to `info`. A name that points into another file, as `'other.xlsm'!Macro`, still goes.

office-defuse decides the `vbaProject.bin` part of an OOXML package, which with `scriptPlugins` set, even to an empty list, goes to the script plugins and not to the file plugins. It decides the project storages of a Word document, a workbook and the other compound files. A project inside an embedded object is for the file plugins that take the object. A PowerPoint 97 project sits compressed in records the rewrite does not change, so it goes as before.

A source that does not decompress, or a dir stream that names a module stream that is not there, leaves nothing to check. The project goes, and MACRO/VBA_PROJECT's `data.reason` says why. The check of the output does not ask the plugins again. It matches the project by the MD5 of its references and module sources, reports the plugins' findings again, and reports MACRO/COMPILED_CODE if the output still holds p-code, which fails the check.

[script-defuse](https://github.com/patrickdk77/script-defuse)'s `formScriptPlugin` is a script plugin for VBA:

```js
const { disarmOffice } = require('@patrickdk77/office-defuse');
const { formScriptPlugin } = require('@patrickdk77/script-defuse');

await disarmOffice(bytes, { scriptPlugins: [formScriptPlugin()] });
```

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
| DUPLICATE_NAME | Two names equal after percent-decoding, Unicode NFKC and full case folding, and without the dots and spaces Windows drops from the end of each segment. A sharp s (U+00DF) meets SS, a ligature meets its letters, and `a.xml.` meets `a.xml` |
| UNSAFE_NAME | An empty name, a name over 1024 bytes, a leading `/`, a drive letter, a backslash, a control character, a segment of only dots and spaces, such as `..`, `.. ` or `...`, which Windows reads as `..` or as empty, or a directory entry with content |
| NAME_ENCODING | A non-ASCII name without the UTF-8 flag, invalid UTF-8, or an Info-ZIP Unicode path field that names another entry |
| ENCRYPTED_ENTRY | Any of the encryption flags |
| UNSUPPORTED_METHOD | Any compression method but stored and deflate |
| ZIP64_MISMATCH | A field at its maximum with no ZIP64 value behind it, a ZIP64 extra field that holds other fields than its header defers to it, a local ZIP64 field that disagrees with its header, an end record that disagrees with the ZIP64 end record, or a locator that does not point at one |
| BAD_EXTRA_FIELD | An extra field block that does not split exactly into records, or a record id given twice |
| DATA_ERROR | Data that does not inflate to its declared size and CRC, deflate data that is corrupt, ends early or leaves compressed bytes unused, or a stored entry with two different sizes |

It reports four more without stopping. LEADING_BYTES is data before the first entry. HIDDEN_DATA is bytes between entries, or before the central directory, that no entry covers. HIDDEN_ENTRY is a local file header the central directory does not list. It can sit in those uncovered bytes, or inside the data of a stored entry that has a data descriptor, since a reader that walks the file finds the end of such an entry only by looking for the next signature. A streaming unzipper extracts a hidden entry as a file, and a reader of the central directory never sees it. TRAILING_BYTES is bytes after the end record that its comment does not cover. Each is `strip`. The rewrite writes no data descriptors, so a reader that walks the copy finds only the listed entries.

The reader inflates every entry once, start to end, and checks it against its CRC and declared size. Inflation stops as soon as the output passes the declared size.

A writer that streams an entry sets the data descriptor flag and may leave the local header's CRC and sizes unset, as zero, or as ones in a local ZIP64 field. The data descriptor after the data then gives them, and they must match the central directory.

The XML tokenizer reads each part in pieces. It refuses these, and stops the run:

- XML/ENCODING: anything but UTF-8 or UTF-16, the two encodings OPC allows, an encoding declaration that disagrees with the bytes, and bytes that are not valid in their encoding. A declaration of us-ascii agrees with UTF-8 bytes as long as every byte is ASCII, since both read them the same. SharePoint writes such declarations into custom XML parts.
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

A run that stops rejects the file with a null score and the UNKNOWN level, whatever `actionOverrides` say about the finding that stopped it. The container refusals, a limit, an unrecognized or ambiguous format, an encrypted file that no password at hand opens or whose algorithm office-defuse does not decrypt, and an unexpected error, PACKAGE/UNREADABLE, all stop a run. The findings made before the stop stay in the report. A time or nesting limit that a plugin or a nested run reaches stops the upload's run too.

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
| `limits.nestingDepth` | none | The deepest a file may sit in the upload. The upload is at 0 and a file inside it at 1, whatever the formats in between |
| `limits.timeMs` | none | The whole run, files inside included |
| `password` | none | One password or a list, tried in order on an encrypted OOXML package or .xls workbook, before VelvetSweatshop |
| `actionOverrides` | none | Change the action for a category, or for a category and detail |
| `scoreWeights`, `scoreBands` | built in | Change the scoring |
| `filePlugins` | none | Plugins that decide about the files inside the file. See [File plugins](#file-plugins) |
| `scriptPlugins` | none | Plugins that decide about the modules of a VBA project. See [Script plugins](#script-plugins) |
| `tempDir` | `os.tmpdir()` | Where temporary files go |
| `memoryThreshold` | 8 MiB | Disarm keeps its copy of the upload, its output and each file it shows a plugin in memory up to this size, and in a temporary file past it. The copies plugins return share this much memory while they wait for the rewrite |

The nine limits with a default bound the memory or work any input can cost, so they apply unless you set them. `Infinity` turns one off. A limit that is reached stops the run with a LIMIT finding.

Besides one XML token, the stack of open elements and the declarations of the part's internal subset, office-defuse holds the central directory's table of entries, which parts each part's relationships reach, and, while it reads a relationships part, the Ids that part has given so far, about 100 bytes each, so it can find one given twice. Content types, manifest paths and object references that name no entry of the archive are not kept. Everything else streams.

## The defuse family

office-defuse is one of a family of packages that each handle one kind of file, beside [pdf-defuse](https://github.com/patrickdk77/pdf-defuse). It declares the family's shared types with the same shapes as pdf-defuse: `DefuseFinding`, `ByteSource`, `ByteSink`, `PluginContext`, `ContainedFile`, `ContainedFileResult` and `ContainedFilePlugin`. It does not depend on pdf-defuse, and TypeScript matches the types by shape.

A finding's `category` and `detail` are plain strings, and `weight`, when set, says how much it adds to a score. The time and nesting limits' errors carry `code: 'DEFUSE_LIMIT'` and `limit: 'time'` or `limit: 'nesting'`, and an I/O error inside a nested run carries `code: 'DEFUSE_IO'` with the error as its `cause`. Every package of the family passes these up the tree of files. At depth 0, office-defuse turns a limit into LIMIT/TIME or LIMIT/NESTING_DEPTH and throws the cause of an I/O error.

`@patrickdk77/file-defuse` is the family's router, one contained-file plugin that every package gets in its `filePlugins`. office-defuse hands it the files it finds inside, with their depth counted from the upload and the upload's deadline, and the router can pass an Office file inside any other file to `officePlugin`.

## Command line

```
office-defuse inspect <file> [options]
office-defuse defuse <in> <out> [options]

--password-file <path>   read passwords from a file, one per line; OFFICE_DEFUSE_PASSWORD holds one
--plugin <module>        load file plugins from a file or an installed package, repeatable
--config <module>        load a whole options object, repeatable; the other options apply on top
--defuse-embedded-office defuse Office files inside the file with office-defuse and keep them
--limit <name>=<value>   fileSize, entries, decompressedBytes, totalDecompressedBytes, xmlDepth,
                         xmlTokenSize, xmlEntityDepth, xmlEntityExpansions, xmlEntitySize, rtfDepth,
                         nestingDepth, timeMs; sizes accept kb, mb, gb
--json                   print the result as JSON
```

Use `-` for stdin or stdout. The tool copies stdin to a temporary file in `tempDir`, which a `--config` module can set. The copy stops one byte past `limits.fileSize`, and the run then rejects the upload without reading the rest. The copy counts toward `limits.timeMs`. `inspect` prints its report to stdout. `defuse` does too, unless the file itself goes to stdout, and then the report goes to stderr. `defuse` refuses an output that is its input under any name, because a failed write deletes the partial output.

`--password-file` reads one password per line. The tool drops a carriage return at the end of a line and skips blank lines. Spaces belong to the password. `OFFICE_DEFUSE_PASSWORD` holds one password, and a line break in it is part of that password. `--password-file` replaces it, and either one replaces a `password` that a `--config` module sets. `--password` is a usage error.

The text report shows each control character as an escape such as `\u001b`. The JSON report escapes DEL and the C1 range too, which JSON itself does not require. A part name, a field code or a link from the file can hold an escape sequence that rewrites what the terminal shows, the verdict included.

`fileSize`, `decompressedBytes` and `totalDecompressedBytes` take sizes. The other limits take whole numbers. Any other value is a usage error.

A `--plugin` module exports `plugins`, an array, or `plugin`, a single plugin, each with `kind: 'file'`. Either can also sit on the default export. ES modules, CommonJS files and installed packages all load, the same way pdf-defuse's tool loads them. A `--config` module's default export is the options object. Several `--config` modules combine. Their plugin lists join in order and their limits merge, and for any other option the later module wins. The flags apply on top of every module, wherever they sit on the command line.

Exit codes are 0 for clean, 1 for defused or strippable, 2 for rejected and 3 for a usage or I/O error. An output or a report that could not be written is an I/O error. An unknown option shows in the error without the text after its `=`.

## What it does not check

- The OLE parts of a ZIP package, such as the objects inside `oleObject*.bin` and ActiveX binaries, and `vbaProject.bin` without script plugins. office-defuse hands them to the file plugins and reads none of them itself.
- VBA source. office-defuse reads the modules only to show them to script plugins, and without plugins reports a project whole. In a user form that stays, the controls' data and pictures in its `o` stream go unread.
- Encrypted Word 6.0 and 95 documents. They are rejected. In a Word document encrypted with RC4 CryptoAPI, the OLE objects that Word keeps in the Data stream rather than in ObjectPool. The pictures of an encrypted presentation.
- Formulas in .xlsb sheets, and in an Excel 5.0 or 95 workbook. A DDE link that a formula uses is found in the workbook's link records.
- Hyperlinks in a Visio .vsd, ShapeSheet cells in a drawing from Visio 5.0 or earlier, and anything in Publisher and Project files but VBA projects and objects.
- An Outlook message's plain text body. What an attached file holds is for the file plugins, and so is the HTML body.
- Images, fonts and other media, and the inside of an embedded Office package such as a chart's workbook. These go to the file plugins, and `officePlugin` reads an embedded package.
- Data connections in an .xlsb, which are binary.
- Word 2003 XML and SpreadsheetML 2003, single XML files with their own vocabulary. They are PACKAGE/UNRECOGNIZED.
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

`test/fixtures/media` holds packages and RTF files with pictures and with the binary parts of the formats themselves, for the file plugins. `python3 -I scripts/make-fixtures.py media` writes them with Python alone.

`test/fixtures/cases` holds one small file for each behavior a run over a corpus of real Office files turned up, written from scratch by `python3 -I scripts/make-fixtures.py cases`, which needs only Python. No corpus file, file name or path is in the repository, macro and OLE payloads are inert markers, and every address uses a reserved name. `manifest.json` gives each file's issue, what office-defuse reports for it, and what the build before the fix it documents reported. A test checks every entry, and that the manifest and the folder list the same files.

## License

MIT
