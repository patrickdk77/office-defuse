import { OfficeCategory as C, OfficeDetail as D, type OfficeActionOverride, type OfficeFinding, type OfficeFindingAction } from './types';

interface Spec {
  action: OfficeFindingAction;
  /**
   * May contain <first|second> choices, which take the first alternative whose {key} values from the data are all
   * set and not empty, or that names none. A value has control and format characters replaced by "?" and is cut to
   * 60 characters.
   */
  description: string;
  weight: number;
}

const specs = new Map<string, Spec>();
function def(c: C, d: D, action: OfficeFindingAction, weight: number, description: string): void {
  specs.set(`${c}/${d}`, { action, weight, description });
}

// CONTAINER: the ZIP layer. Every reject stops the run.
def(C.Container, D.NotZip, 'reject', 0, 'The file is not a ZIP archive, or it is cut short');
def(C.Container, D.AmbiguousEnd, 'reject', 0, 'The archive has more than one end record a reader could pick');
def(C.Container, D.MultiDisk, 'reject', 0, 'The archive claims to span several disks');
def(C.Container, D.DirectoryMismatch, 'reject', 0, "The archive's end record disagrees with its central directory");
def(C.Container, D.OutOfRange, 'reject', 0, 'An entry, a sector or a directory lies outside the file');
def(C.Container, D.Overlap, 'reject', 0, 'Two entries or streams of the file share bytes');
def(C.Container, D.HeaderMismatch, 'reject', 0, 'A local header contradicts the central directory<: {name}|>');
def(C.Container, D.DuplicateName, 'reject', 0, 'Two entries have the same name, ignoring case<: {name}|>');
def(C.Container, D.UnsafeName, 'reject', 0, 'An entry name is unsafe to use as a path<: {name}|>');
def(C.Container, D.NameEncoding, 'reject', 0, 'An entry name can be read in more than one way');
def(C.Container, D.EncryptedEntry, 'reject', 0, 'An entry is encrypted<: {name}|>');
def(C.Container, D.UnsupportedMethod, 'reject', 0, 'An entry uses a compression method other than stored or deflate<: {name}|>');
def(C.Container, D.Zip64Mismatch, 'reject', 0, 'ZIP64 fields disagree with the fields they extend');
def(C.Container, D.BadExtraField, 'reject', 0, 'An extra field block does not parse<: {name}|>');
def(C.Container, D.DataError, 'reject', 0, "An entry's data does not match its declared size or checksum<: {name}|>");
def(C.Container, D.LeadingBytes, 'strip', 10, 'Data before the first entry of the archive');
def(C.Container, D.HiddenData, 'strip', 20, 'Data in the file that no entry or stream covers');
def(C.Container, D.HiddenEntry, 'strip', 40, 'A local file header the central directory does not list, which a reader that walks the file in order extracts<: at offset {offset}|>');
def(C.Container, D.TrailingBytes, 'strip', 10, 'Data after the end of the archive or compound file');
def(C.Container, D.BadHeader, 'reject', 0, 'The compound file header is invalid<: {reason}|>');
def(C.Container, D.BrokenChain, 'reject', 0, 'A sector chain of the compound file runs into a free or reserved sector');
def(C.Container, D.ChainLoop, 'reject', 0, 'A sector chain of the compound file loops');
def(C.Container, D.DirectoryLoop, 'reject', 0, 'The compound file directory reaches one entry twice');
def(C.Container, D.BadEntry, 'reject', 0, 'A compound file directory entry is invalid<: {reason}|>');
// XML. MALFORMED and ENCODING stop the run.
def(C.Xml, D.Doctype, 'info', 0, 'An XML part declares a document type<: {part}|>');
def(C.Xml, D.InternalSubset, 'strip', 20, 'An XML part declares entities or other markup in its document type<: {part}|>');
def(C.Xml, D.ExternalEntity, 'strip', 40, 'An XML part takes content from <{target}|outside the file>');
def(C.Xml, D.Malformed, 'reject', 0, 'An XML part is not well-formed<: {part}|>');
def(C.Xml, D.Encoding, 'reject', 0, 'An XML part is not in valid UTF-8 or UTF-16<: {part}|>');
// LIMIT
def(C.Limit, D.FileSize, 'reject', 0, 'The file exceeds the file size limit');
def(C.Limit, D.EntryCount, 'reject', 0, 'The file exceeds the entry count limit');
def(C.Limit, D.DecompressedSize, 'reject', 0, 'An entry exceeds the decompressed size limit');
def(C.Limit, D.TotalDecompressedSize, 'reject', 0, 'The file exceeds the total decompressed size limit');
def(C.Limit, D.XmlDepth, 'reject', 0, 'An XML part exceeds the element depth limit');
def(C.Limit, D.XmlTokenSize, 'reject', 0, 'An XML part exceeds the token size limit');
def(C.Limit, D.XmlEntityDepth, 'reject', 0, 'An XML part exceeds the entity nesting limit');
def(C.Limit, D.XmlEntityExpansions, 'reject', 0, 'The file exceeds the entity expansion limit');
def(C.Limit, D.XmlEntitySize, 'reject', 0, 'The file exceeds the expanded entity size limit');
def(C.Limit, D.RtfDepth, 'reject', 0, 'The RTF file exceeds the group nesting limit');
def(C.Limit, D.NestingDepth, 'reject', 0, 'The file sits deeper in the upload than the nesting limit allows');
def(C.Limit, D.Time, 'reject', 0, 'The file exceeds the time limit');
// PACKAGE
def(C.Package, D.Unrecognized, 'reject', 0, 'The file is not a recognized Office or OpenDocument file');
def(C.Package, D.AmbiguousFormat, 'reject', 0, 'The file can be read as more than one format');
def(C.Package, D.Unreadable, 'reject', 0, 'The file could not be read');
def(C.Package, D.UnreadableBody, 'strip', 10, "An Outlook message's RTF body does not decompress<: {stream}|>");
def(C.Package, D.DuplicateEntry, 'reject', 0, 'The package names one thing twice, and readers may pick either');
def(C.Package, D.UntypedPart, 'strip', 10, 'A part has no content type<: {part}|>');
def(C.Package, D.UnreferencedPart, 'strip', 0, 'A part nothing in the package refers to<: {part}|>');
def(C.Package, D.OdfLayout, 'strip', 10, 'The OpenDocument package layout is irregular<: {reason}|>');
// ENCRYPTED
def(C.Encrypted, D.PasswordRequired, 'reject', 0, 'The document is encrypted with a password');
def(C.Encrypted, D.DefaultPassword, 'strip', 10, 'The file is encrypted with <{password}|VelvetSweatshop>, the password <{application}|Excel> opens it with without asking');
def(C.Encrypted, D.UserPassword, 'strip', 0, 'The file was decrypted with the supplied password');
def(C.Encrypted, D.Rc4_40, 'info', 0, 'Encryption is RC4 with a 40-bit key');
def(C.Encrypted, D.Rc4_128, 'info', 0, 'Encryption is RC4 with a 128-bit key');
def(C.Encrypted, D.Rc4Other, 'info', 0, 'Encryption is RC4 with a <{bits}|longer>-bit key');
def(C.Encrypted, D.Aes128, 'info', 0, 'Encryption is AES with a 128-bit key');
def(C.Encrypted, D.Aes192, 'info', 0, 'Encryption is AES with a 192-bit key');
def(C.Encrypted, D.Aes256, 'info', 0, 'Encryption is AES with a 256-bit key');
def(C.Encrypted, D.XorObfuscation, 'info', 0, 'Encryption is XOR obfuscation');
def(C.Encrypted, D.UnsupportedAlgorithm, 'reject', 0, 'The file is encrypted with an algorithm office-defuse does not decrypt<: {algorithm}|>');
// MACRO
def(C.Macro, D.VbaProject, 'strip', 70, 'The document holds VBA macros');
def(C.Macro, D.CompiledCode, 'strip', 10, "The VBA project's compiled code was removed, so Office compiles the source the plugins checked");
def(C.Macro, D.PluginPassed, 'info', 30, 'A VBA module was kept by <plugin {plugin}|a plugin>');
def(C.Macro, D.PluginRemoved, 'strip', 0, '<Plugin {plugin}|A plugin> did not keep a VBA module, so the VBA project was removed');
def(C.Macro, D.PluginFailed, 'strip', 0, '<Plugin {plugin}|A plugin> failed on a VBA module, so the VBA project was removed');
def(C.Macro, D.XlmMacrosheet, 'strip', 70, 'The workbook holds an Excel 4.0 macro sheet');
def(C.Macro, D.WordBasic, 'strip', 70, 'The Word 6.0 or 95 document holds WordBasic macros, which Word runs as VBA');
def(C.Macro, D.OdfBasic, 'strip', 70, 'The document holds Basic macros');
def(C.Macro, D.OdfScripts, 'strip', 70, 'The document holds scripts');
def(C.Macro, D.OdfDialogs, 'strip', 30, 'The document holds Basic dialogs, which call macros');
def(C.Macro, D.EventOnOpen, 'strip', 70, 'A macro or script runs when the document opens');
def(C.Macro, D.Event, 'strip', 50, 'A macro or script runs on <the {event}|an> event');
def(C.Macro, D.Link, 'strip', 50, 'A link or action runs a macro or script');
// UI
def(C.Ui, D.Ribbon, 'info', 0, 'The document customizes the ribbon');
def(C.Ui, D.Callback, 'strip', 30, 'The custom ribbon calls a macro<: {callback}|>');
def(C.Ui, D.OdfConfiguration, 'strip', 10, 'The document carries its own menu, toolbar or key settings, which can call macros');
// ACTION
def(C.Action, D.RunProgram, 'strip', 80, 'A click or hover runs a program or opens a file');
def(C.Action, D.SubmitForm, 'strip', 40, 'A form sends its data to a URL<: {target}|>');
def(C.Action, D.OleVerb, 'strip', 40, 'A click or hover activates an embedded object');
def(C.Action, D.ObjectUpdate, 'strip', 50, 'An embedded object updates itself as the document opens, which starts the program it belongs to');
// DDE
def(C.Dde, D.Field, 'strip', 70, 'A field runs a command through DDE');
def(C.Dde, D.Link, 'strip', 70, 'A DDE link asks another program for data<: {application}|>');
def(C.Dde, D.Formula, 'strip', 70, 'A formula <takes text from external entity {entity}, so it may call DDE|asks another program for data through DDE>');
// FIELD
def(C.Field, D.Macrobutton, 'strip', 40, 'A field runs a macro when double-clicked');
def(C.Field, D.Computed, 'strip', 50, "A field's type <comes from external entity {entity}|is built by another field when it updates>");
// EXTERNAL
def(C.External, D.Template, 'strip', 60, 'The document loads its template from <{target}|outside the file>');
def(C.External, D.Object, 'strip', 60, 'A linked object loads from <{target}|outside the file>');
def(C.External, D.Image, 'strip', 30, 'An image loads from <{target}|outside the file>');
def(C.External, D.Frame, 'strip', 40, 'A frame shows <{target}|a page from outside the file>');
def(C.External, D.Include, 'strip', 40, 'Content is included from <{target}|another file>');
def(C.External, D.Workbook, 'strip', 30, 'Formulas read another workbook at <{target}|another location>');
def(C.External, D.Reload, 'strip', 40, 'The document loads <{target}|another address> on its own after a delay');
def(C.External, D.Other, 'strip', 30, 'The document refers to <{target}|something outside the file>');
def(C.External, D.DataConnection, 'strip', 40, 'A data connection loads from outside the file<: {target}|>');
def(C.External, D.LocalFile, 'strip', 5, 'The document refers to a local file<: {target}|>');
// LINK
def(C.Link, D.Safe, 'info', 0, 'Web or email link');
def(C.Link, D.FileUrl, 'strip', 60, 'A link opens a local file');
def(C.Link, D.NetworkPath, 'strip', 60, 'A link opens a network share');
def(C.Link, D.DataUrl, 'strip', 60, 'A link carries embedded data');
def(C.Link, D.ScriptUrl, 'strip', 50, 'A link uses a javascript: or vbscript: URL');
def(C.Link, D.OtherScheme, 'strip', 30, 'A link uses a scheme other than http, https or mailto');
def(C.Link, D.Credentials, 'strip', 40, 'A link hides its real host behind a user name');
def(C.Link, D.IpHost, 'strip', 40, 'A link points at a raw IP address');
def(C.Link, D.LookalikeHost, 'strip', 40, "A link's host uses punycode or mixed scripts");
def(C.Link, D.EncodedHost, 'strip', 40, "A link's host is percent-encoded");
def(C.Link, D.Relative, 'strip', 30, "A relative link, which resolves against the file's own location");
def(C.Link, D.TextMismatch, 'strip', 40, "A link's text or tooltip names a different site than it opens");
def(C.External, D.Formula, 'strip', 40, 'A formula fetches from outside the file when it is calculated<: {function}|>');
def(C.Macro, D.Formula, 'strip', 60, 'A formula runs code outside the file when it is calculated<: {function}|>');
// EMBEDDED
def(C.Embedded, D.OleObject, 'strip', 40, 'An embedded OLE object<, which holds the file {fileName}|>');
def(C.Embedded, D.Package, 'strip', 30, 'An embedded document<: {contentType}|>');
def(C.Embedded, D.OleFile, 'strip', 40, 'A part holds an OLE file that nothing declares as an object');
def(C.Embedded, D.Applet, 'strip', 60, 'An embedded Java applet');
def(C.Embedded, D.Plugin, 'strip', 50, 'An embedded browser plug-in object');
def(C.Embedded, D.Equation, 'strip', 70, 'An Equation Editor 3.0 object, which CVE-2017-11882 and CVE-2018-0802 exploit');
def(C.Embedded, D.Attachment, 'strip', 30, '<An attached file "{name}"|An attached file> (<{type}|unknown type>)');
def(C.Embedded, D.PluginPassed, 'info', 20, 'An embedded file was kept by <plugin {plugin}|a plugin> with nothing removed');
def(C.Embedded, D.PluginScrubbed, 'info', 10, 'An embedded file was cleaned and kept by <plugin {plugin}|a plugin>');
def(C.Embedded, D.PluginFailed, 'strip', 0, '<Plugin {plugin}|A plugin> failed, so the embedded file was removed');
def(C.Embedded, D.PluginRemoved, 'strip', 0, 'An embedded file was removed by <plugin {plugin}|a plugin>');
// CONTROL
def(C.Control, D.ActiveX, 'strip', 50, 'An ActiveX control<: {reason}|>');
def(C.Control, D.FormsControl, 'info', 10, 'A Forms 2.0 control<, {className}|>, whose persisted properties office-defuse read');
def(C.Control, D.ControlPicture, 'strip', 10, "A control's picture for printing, a Windows metafile");
// PROCESSING
def(C.Processing, D.VerificationFailed, 'reject', 0, 'The defused output did not pass its own inspection, so nothing was written');

/** The built-in spec for a kind of finding, or undefined for a kind office-defuse does not make, such as one from another package. */
export function findingSpec(category: string, detail: string): Spec | undefined {
  return specs.get(`${category}/${detail}`);
}

/** The spec of a kind office-defuse makes itself, which always has one. */
function ownSpec(category: C, detail: D): Spec {
  const s = findingSpec(category, detail);
  if (!s) throw new Error(`No finding spec for ${category}/${detail}`);
  return s;
}

/** Every category and detail pair office-defuse makes, with its default action, description and score weight. */
export function allFindingSpecs(): Array<{ category: C; detail: D } & Spec> {
  return [...specs.entries()].map(([k, s]) => {
    const [category, detail] = k.split('/') as [C, D];
    return { category, detail, ...s };
  });
}

/** Values from the file can hold control or direction characters, so the description shows none of them. */
const shown = (v: string | number) => String(v).replace(/[\p{Cc}\p{Cf}]/gu, '?');

export class FindingFactory {
  constructor(private readonly overrides: OfficeActionOverride[] = []) {}

  /** The action of a kind of finding, after the caller's overrides. */
  actionFor(category: C, detail: D): OfficeFindingAction {
    const exact = this.overrides.find(o => o.category === category && o.detail === detail);
    if (exact) return exact.action;
    const byCategory = this.overrides.find(o => o.category === category && o.detail === undefined);
    if (byCategory) return byCategory.action;
    return ownSpec(category, detail).action;
  }

  make(category: C, detail: D, location?: string, data?: Record<string, string | number>): OfficeFinding {
    const spec = ownSpec(category, detail);
    const description = spec.description.replace(/<([^<>]*)>/g, (_match, body: string) => {
      const fills = body.split('|').map(alt => {
        let missing = false;
        const text = alt.replace(/\{([a-zA-Z]+)\}/g, (_, key: string) => {
          const v = data?.[key];
          if (v === undefined || v === '') missing = true;
          const s = shown(v ?? '');
          return s.length > 60 ? `${s.slice(0, 57)}...` : s;
        });
        return { text, missing };
      });
      // Every choice ends with an alternative that names no value, so one always fits.
      return (fills.find(x => !x.missing) ?? fills[fills.length - 1]).text;
    });
    const f: OfficeFinding = { category, detail, description, action: this.actionFor(category, detail) };
    if (location !== undefined) f.location = location;
    if (data !== undefined && Object.keys(data).length) f.data = data;
    return f;
  }
}
