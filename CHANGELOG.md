# Changelog

## Unreleased

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
