# Changelog

## 0.3.0 - 2026-10-17

- A plugin's answer is final. A file a plugin passed, scrubbed or removed gets none of office-defuse's own findings, such as EMBEDDED/PACKAGE, EMBEDDED/OLE_OBJECT, EMBEDDED/EQUATION or MACRO/VBA_PROJECT, and office-defuse no longer looks inside it for VBA projects and objects. 0.2.4 reported those findings next to the plugin's, scored a removed file at office-defuse's weight, and removed an Equation Editor object a plugin passed.
- A VBA project whose modules the script plugins all passed or removed is reported by MACRO/PLUGIN_PASSED and PLUGIN_REMOVED and their findings, without MACRO/VBA_PROJECT. A project the plugins removed scored 70 for MACRO/VBA_PROJECT on top of what they found.
- A button, shape, callback, field or action that runs a macro of a project the plugins answered for weighs 0, kept or removed, in the run's score and in the weights a container gets through `officePlugin`. A container counted the kept ones at 50 before, and a removed one counted in every score.
- MACRO/PLUGIN_PASSED and EMBEDDED/PLUGIN_PASSED and PLUGIN_SCRUBBED weigh 0, in the upload's score and the output's. The output's score counted them before.

## 0.2.4 - 2026-10-16

- Forms 2.0 controls stay: CommandButton, Label, TextBox, ListBox, ComboBox, CheckBox, OptionButton, ToggleButton, SpinButton and ScrollBar, when their persisted properties read as Office writes them and hold no picture. The new CONTROL/FORMS_CONTROL, `info`, reports each. It covers OOXML activeX parts persisted as a stream, a storage or a property bag, Word's control storages and Excel's `Ctls` stream. Every other control is CONTROL/ACTIVEX as before, now with `data.reason`.
- A Word control's picture for printing, its `\x03PRINT` metafile, is the new CONTROL/CONTROL_PICTURE, `strip`, and goes while the control stays.
- Excel workbooks are written again record by record instead of rejected. A macro sheet becomes an empty worksheet, a name that runs a macro is renamed, a button or shape loses its macro, a hyperlink goes, and an external workbook, DDE or OLE link points at the workbook itself. The stream positions in BoundSheet8, Index, DBCell and ExtSST follow. Excel 5.0 and 95 workbooks get the sheet and name edits.
- PowerPoint 97 to 2003 presentations are written again as a full save. Objects, controls and VBA projects that go take the records that name them, actions that run a program, a macro or an OLE verb do nothing, unsafe hyperlinks go, and linked media lose their path. A VBA project in a presentation goes to the script plugins like any other.
- An encrypted workbook or presentation that is written again, as one VelvetSweatshop opens is, comes out decrypted. A presentation with a Pictures or EncryptedSummary stream, which is encrypted too, is not written again.
- .ppa add-ins are read as PowerPoint files. 0.2.3 called them PACKAGE/UNRECOGNIZED.
- A button macro's formula in an Excel Obj record is read as Excel writes it, with no size before the token count.

## 0.2.3 - 2026-10-16

- `scriptPlugins` shows each module of a VBA project to script plugins, such as script-defuse's `formScriptPlugin`. A project whose modules all pass stays, written without its compiled code, so Office compiles the source the plugins read and stomped p-code never reaches the output. The `vbaProject.bin` of an OOXML package and the project storages of a compound file go to the plugins. A project inside an embedded object or a PowerPoint 97 presentation goes as before.
- New findings: MACRO/PLUGIN_PASSED, MACRO/PLUGIN_REMOVED, MACRO/PLUGIN_FAILED and MACRO/COMPILED_CODE.
- A kept project keeps its user forms when they hold Forms 2.0 controls only. A form whose class table could name another ActiveX control gets CONTROL/ACTIVEX, and the project goes.
- An Excel shape, form control, ribbon callback, PowerPoint action or MACROBUTTON field that runs one of a kept project's macros by name stays, and its finding is `info`.
- The output check matches a kept project by its sources, and fails when the output still holds p-code.
- The VBA reader takes a chunk header without its signature bits, and a few bytes of padding after the last chunk. Office opens files written that way, and the output holds neither, since a kept module is compressed again.

## 0.2.2 - 2026-10-16

- An add-in's empty presentation part is read and written as it is. PowerPoint writes every `.ppam` that way, since an add-in has no slides, and 0.2.1 rejected each one as XML/MALFORMED. An empty main part in any other format is still rejected.
- A part that declares us-ascii over bytes that are all ASCII is read as UTF-8, which it is. SharePoint writes such parts into custom XML, and 0.2.1 rejected the whole file as XML/ENCODING. A us-ascii declaration over any byte outside ASCII is still XML/ENCODING.

## 0.2.1 - 2026-10-11

- `detectOffice` and `detectOfficeSource` tell a router whether office-defuse reads a file. They name its format and MIME type with the functions the inspection names them with, and claim nothing office-defuse cannot name, such as a plain ZIP archive, a NuGet package or an installer. An encrypted OOXML package that no password at hand opens is claimed as a compound file, since only the password names the format inside. `officePlugin` accepts exactly what they claim, so it no longer takes a bare OLE object. The inspection now names the format before it reads on, so an encrypted Word or Excel template is `dot` or `xlt` without its password, and a Flat OPC package rejected after that point reports `xml`.

## 0.2.0 - 2026-10-11

### Disarm

- `disarmOffice` and `disarmOfficeSource` write a defused copy. The result has pdf-defuse's shape, a `status` of `clean`, `defused` or `rejected` with `before`, `after` and `removed`. The copy leaves out what each strip finding names, and reaches the sink only after an inspection of its own finds nothing but `info` findings. Otherwise the upload is rejected with the new PROCESSING/VERIFICATION_FAILED.
- In OOXML and OpenDocument packages, a part goes with every reference to it, in relationships, `[Content_Types].xml`, the manifest and the elements that name it. Word fields lose their code and keep their result, hyperlinks keep their text, and an Excel 4.0 macro sheet becomes an empty worksheet. An internal subset goes with its entity references written out, and references to external entities go. An encrypted package comes out decrypted.
- In compound files, the output leaves out VBA projects, object and control storages, Excel's `Ctls` stream and Outlook attachments, and drops hidden data. Content inside a stream, such as Word fields or Excel records, stays, so a file with findings there is rejected.
- In RTF, the output leaves out field instructions, objects, `\objupdate`, the template, shape hyperlinks and text after the document. An object's result group takes its place.
- `bufferSink`, `fileSink` and `writableSink` are exported for `disarmOfficeSource`.
- The copy keeps other formats' signatures out of the places their readers look. A compound file's FAT comes right after the header, so no stream sits in the first 1024 bytes, where PDF readers look for a header. Free sectors of zeros keep an embedded package's ZIP end record out of the last 65,557 bytes, where ZIP readers look for one. A ZIP entry whose name, local header or compressed bytes would put a PDF header in the first 1024 bytes is deflated again, or written after the entries that follow it.

### ZIP reader

- A local file header the central directory does not list is the new CONTAINER/HIDDEN_ENTRY, `strip`. It sits between entries, or inside a stored entry with a data descriptor. A streaming unzipper extracts it as a file, while a reader of the central directory never sees it.
- Entry names collide after NFKC and full case folding. A name with a sharp s (U+00DF) and the same name with SS are one file to macOS's case-insensitive file system, and now CONTAINER/DUPLICATE_NAME here. So are a ligature and its letters, and two names that differ only by dots and spaces at the end of a segment, which Windows drops.
- A path segment of only dots and spaces, such as `.. ` or `...`, is CONTAINER/UNSAFE_NAME as `..` is, since Windows drops the trailing dots and spaces.

### File plugins

- `filePlugins` takes pdf-defuse's contained-file plugins. They see OLE parts, embedded packages, pictures and other whole files in packages, Outlook attachments, embedded objects and Packager files in compound files, and objects and PNG, JPEG and EMF pictures in RTF. A plugin's verdict decides whether a file stays and what bytes it keeps. A file no plugin accepts gets the rule it got in 0.1.1.
- New findings EMBEDDED/PLUGIN_PASSED, PLUGIN_SCRUBBED, PLUGIN_FAILED and PLUGIN_REMOVED report each verdict.
- `officePlugin(options)` lets another package, or office-defuse itself, defuse an Office file inside another file. It accepts by content alone.
- `limits.nestingDepth` bounds how deep a file may sit in the upload, as LIMIT/NESTING_DEPTH. Time and nesting limits from plugins, and their I/O errors, stop the whole upload.

### Command-line tool

- `office-defuse inspect <file>` and `office-defuse defuse <in> <out>` take `-` for stdin and stdout, and `--password-file`, `--plugin`, `--config`, `--defuse-embedded-office`, `--limit` and `--json`. They exit with pdf-defuse's codes: 0 clean, 1 defused or strippable, 2 rejected, 3 usage or I/O error.
- `--password-file` holds one password per line, and `OFFICE_DEFUSE_PASSWORD` holds one. A password given as an argument is a usage error.

### Formats

- XPS and OpenXPS packages, .xps and .oxps. The markup's references to documents, pages, fonts, images and color profiles count as reaching their parts, and those outside the package are EXTERNAL. Links with a scheme get the link checks, and pictures go to the file plugins.
- Flat OpenDocument files, .fodt .fods .fodp .fodg, get the OpenDocument rules. Every relative reference in one names a file beside the document. Pictures and OLE objects held as base64 go to the file plugins.
- Flat OPC, the XML form of Word, Excel and PowerPoint files, is read as the package of its parts and written back as one XML file.
- A chart's `xlink:href=".."` and a `vnd.sun.star.Package:` URL name the document itself and its own storage, so they no longer count as EXTERNAL.

### Links

- A safe link whose text, tooltip or display name is an address of a different site than it opens is the new LINK/TEXT_MISMATCH. Sites compare as in pdf-defuse. That covers Word hyperlinks and HYPERLINK fields in .docx, .doc and RTF, DrawingML runs, Excel hyperlinks, OpenDocument links and XPS glyphs. The rewrite removes the link and keeps its text.

### Formulas, connections and forms

- Formulas that call WEBSERVICE or IMAGE are the new EXTERNAL/FORMULA, and those that call RTD, CALL, REGISTER, REGISTER.ID or EXEC the new MACRO/FORMULA, in .xlsx, .xls and OpenDocument formulas. The rewrite removes the formula and keeps the cell's value.
- The sources of Excel data connections are the new EXTERNAL/DATA_CONNECTION, and the rewrite empties them.
- An OpenDocument form's submit target is the new ACTION/SUBMIT_FORM. Word's hyperlink base gets the link checks.

### Word 6.0 and 95

- A file saved fast is read in the order its piece table gives, so a field split over pieces stored apart reads whole.

### Visio

- A .vsd from Visio 2000 to 2010 gets the ShapeSheet check a .vsdx gets. A formula that calls RUNADDON, RUNADDONWARGS, CALLTHIS, RUNMACRO, OPENFILE or DOOLEVERB is MACRO/EVENT in an event cell and MACRO/LINK in any other. A VisioDocument stream whose structure does not read is PACKAGE/UNREADABLE. Disarm rejects a drawing with such a cell, since it copies the stream as it is.

### Outlook messages

- A message's RTF body is decompressed and gets the RTF rules, as do the bodies of embedded messages. The rewrite removes what they strip and stores the body uncompressed. A body that does not decompress is the new PACKAGE/UNREADABLE_BODY and goes, leaving the plain text body.
- The RTF and HTML bodies go to the file plugins.

### Encryption

- Word 97 to 2003 documents encrypted with RC4 or RC4 CryptoAPI, or obfuscated with XOR, are decrypted and get the Word rules. They used to be rejected.
- PowerPoint 97 to 2003 presentations encrypted with RC4 CryptoAPI are decrypted. PowerPoint's own default password, `/01Hannes Ruescher/01`, is tried after the caller's, as VelvetSweatshop is for Excel.
- Excel workbooks obfuscated with XOR are decrypted, and the new ENCRYPTED/XOR_OBFUSCATION names the method. They used to be UNSUPPORTED_ALGORITHM.
- Agile encryption in CFB chaining is decrypted, built on AES-ECB.
- ENCRYPTED/DEFAULT_PASSWORD names the password and the application that tries it.

### Options

- `password` takes a list. Each password is tried in order, each once, before VelvetSweatshop.
- `tempDir` and `memoryThreshold` bound the memory disarm uses for its copies, as in pdf-defuse.

## 0.1.1 - 2026-10-11

### New formats

- Compound files: Word 6.0 to 2003 .doc and .dot, Excel 5.0 to 2003 .xls, .xlt and .xla, PowerPoint 97 to 2003 .ppt, Visio .vsd, and Outlook .msg and .oft. Publisher .pub and Project .mpp get the checks every compound file gets, for VBA projects and objects. 0.1.0 rejected all of them as CONTAINER/NOT_ZIP or DIRECTORY_MISMATCH.
- RTF, with its objects, fields, template and shape hyperlinks. 0.1.0 rejected it as CONTAINER/NOT_ZIP.

### Encryption

- Encrypted OOXML packages and RC4-encrypted .xls workbooks open with the new `password` option or with VelvetSweatshop, Excel's built-in password, and get every check of their format. Any other encrypted file is ENCRYPTED/PASSWORD_REQUIRED or ENCRYPTED/UNSUPPORTED_ALGORITHM.

### Findings

- New details: CONTAINER/BAD_HEADER, BROKEN_CHAIN, CHAIN_LOOP, DIRECTORY_LOOP and BAD_ENTRY, LIMIT/RTF_DEPTH, ENCRYPTED/DEFAULT_PASSWORD, USER_PASSWORD, UNSUPPORTED_ALGORITHM and one detail per cipher, MACRO/WORDBASIC, ACTION/OBJECT_UPDATE, EMBEDDED/EQUATION and EMBEDDED/ATTACHMENT.
- In ZIP workbooks too, a defined name that starts with Auto_Open or Auto_Activate is MACRO/EVENT_ON_OPEN, and one that starts with Auto_Close or Auto_Deactivate is MACRO/EVENT.
- Container and limit findings that compound files share with ZIP archives describe both.

### Options

- `limits.rtfDepth` bounds RTF group nesting, 256 by default.

### Package

- `engines` asks for Node 18 instead of 22. Nothing in the package needed 22.
- ZIP CRCs use zlib.crc32 where Node has it, about 50 times faster than the lookup table. Node 18 and early releases of Node 20 and 22 lack it and keep the table.

## 0.1.0 - 2026-10-10

First release.
