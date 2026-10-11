import { TextDecoder } from 'node:util';
import { StopError } from './errors';
import { OfficeCategory as C, OfficeDetail as D } from './types';

export const XML_NS = 'http://www.w3.org/XML/1998/namespace';
export const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';

/**
 * Stands in text for a reference to an external entity, whose content office-defuse never fetches: UNKNOWN, the
 * entity's name, UNKNOWN again. U+FFFF cannot occur in XML, so nothing a part holds is taken for it.
 */
export const UNKNOWN = String.fromCharCode(0xffff);
/** One reference to an external entity, with the entity's name as its group. Built from UNKNOWN so the source stays ASCII. */
const UNKNOWN_NAME = `${UNKNOWN}([^${UNKNOWN}]*)${UNKNOWN}`;
const UNKNOWN_REFERENCE = new RegExp(UNKNOWN_NAME, 'g');
const UNKNOWN_FIRST = new RegExp(UNKNOWN_NAME);

/** A value with each reference to an external entity written back as `&name;`, for a finding to show. */
export const showReferences = (s: string) => (s.includes(UNKNOWN) ? s.replace(UNKNOWN_REFERENCE, '&$1;') : s);

/** The name of the first external entity whose content stands in `s`, or undefined when none does. */
export const unknownEntity = (s: string) => UNKNOWN_FIRST.exec(s)?.[1];

export interface XmlAttribute {
  /** Namespace URI, '' for none. */
  ns: string;
  local: string;
  qname: string;
  /** Normalized value, with references expanded. An attribute-list declaration can give it a default or normalize it further. */
  value: string;
}

export interface XmlElement {
  /** Namespace URI, '' for none. */
  ns: string;
  local: string;
  qname: string;
  /** Attributes other than namespace declarations. */
  attributes: XmlAttribute[];
  /** 1 for the root element. */
  depth: number;
  /**
   * Offsets of the start tag in the decoded text, end exclusive. An element from an entity's text has the offsets of
   * the reference to that entity in the part, or of the outermost one when references nest.
   */
  start: number;
  end: number;
  /** The element's position among the elements of the part, from 0, counting those in entity text. */
  index: number;
  /** The namespace declarations of the start tag, as prefix ('' for the default namespace) and URI, defaults included. */
  namespaces: Array<[string, string]>;
  /** Written as an empty-element tag. */
  empty: boolean;
  /**
   * The start tag does not say on its own what the attributes are: a value refers to an entity the internal subset
   * declares, or an attribute-list declaration added a default or normalized a value.
   */
  rewritten: boolean;
}

export interface XmlHandler {
  start?(el: XmlElement): void;
  /** `start` and `end` are the offsets of the end tag, or of the start tag again for an empty element. */
  end?(el: XmlElement, start: number, end: number): void;
  /** Character data and CDATA content, decoded, in pieces of any size. A reference to an external entity comes as one piece. */
  text?(text: string): void;
}

export interface XmlLimits {
  depth: number;
  tokenSize: number;
  /** Entity references nested inside entity text. */
  entityDepth: number;
  /** Entity references expanded, counted in the budget. */
  entityExpansions: number;
  /** Characters of entity text expanded, counted in the budget. */
  entitySize: number;
}

/** Entity expansion so far. One inspection shares it among its parts, so the limits bound the whole package. */
export interface EntityBudget {
  expansions: number;
  size: number;
}

/** A document type declaration. */
export interface XmlDoctype {
  /** The root element type it names. */
  name: string;
  publicId?: string;
  systemId?: string;
  /** Whether it has an internal subset, declarations between square brackets. */
  internalSubset: boolean;
  /** General entities the internal subset declares, and how many of them are external. */
  entities: number;
  externalEntities: number;
  /** Attribute defaults the internal subset declares. */
  attributeDefaults: number;
  /** Entity and attribute-list declarations skipped after a parameter entity that was not read. */
  skipped: number;
  /** Offsets of the declaration in the decoded text, end exclusive. */
  start: number;
  end: number;
}

/** A reference to an external entity in content, which is never fetched. */
export interface XmlExternalReference {
  name: string;
  publicId?: string;
  systemId: string;
  /** Offsets of the reference in the decoded text, or of the outermost reference whose entity text holds it. */
  start: number;
  end: number;
}

/** What the tokenizer reports besides elements and text. */
export interface XmlNotes {
  doctype?(d: XmlDoctype): void;
  externalEntity?(r: XmlExternalReference): void;
  /**
   * A reference to an internal entity in content, outside any other entity's text, at these offsets. The events of
   * its text come between this call and `referenceEnd`.
   */
  referenceStart?(start: number, end: number): void;
  referenceEnd?(): void;
}

export interface XmlOptions {
  /** Reads the dialect Office writes in VML drawing parts. See XmlTokenizer. */
  vml?: boolean;
  notes?: XmlNotes;
  /** Shared with the other parts of a package. A tokenizer without one counts on its own. */
  budget?: EntityBudget;
}

/** The value of the attribute with this namespace and local name. */
export function attr(el: XmlElement, ns: string, local: string): string | undefined {
  for (const a of el.attributes) if (a.local === local && a.ns === ns) return a.value;
  return undefined;
}

// XML 1.0 fifth edition NameStartChar and the further NameChar ranges, without the colon, which namespaces reserve.
const NAME_START =
  'A-Z_a-z\\u{C0}-\\u{D6}\\u{D8}-\\u{F6}\\u{F8}-\\u{2FF}\\u{370}-\\u{37D}\\u{37F}-\\u{1FFF}\\u{200C}-\\u{200D}\\u{2070}-\\u{218F}\\u{2C00}-\\u{2FEF}\\u{3001}-\\u{D7FF}\\u{F900}-\\u{FDCF}\\u{FDF0}-\\u{FFFD}\\u{10000}-\\u{EFFFF}';
const NAME_MORE = '\\-.0-9\\u{B7}\\u{300}-\\u{36F}\\u{203F}-\\u{2040}';
const NC_NAME = new RegExp(`^[${NAME_START}][${NAME_START}${NAME_MORE}]*$`, 'u');
/** A name token, as an enumerated attribute type lists them. */
const NMTOKEN = new RegExp(`^[${NAME_START}${NAME_MORE}:]+$`, 'u');

/** An XML name without a colon. */
export const isNcName = (s: string) => NC_NAME.test(s);

/** A name with at most one colon, between two names without one. */
const isQName = (s: string) => {
  const colon = s.indexOf(':');
  return colon < 0 ? isNcName(s) : isNcName(s.slice(0, colon)) && isNcName(s.slice(colon + 1));
};

const PREDEFINED: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
/** Longest reference read: "&#x" with eight digits and ";" is 12, and leading zeros are legal. */
const MAX_REFERENCE = 64;
/** The characters a public identifier may hold. */
const PUBID = /^[-a-zA-Z0-9 \r\n'()+,./:=?;!*#@$_%]*$/;
/** Characters that end a name inside a document type declaration. */
const NAME_END = /[\s>[\]"'%;&<]/;
/** Attribute types other than CDATA, whose values are normalized further. NOTATION and enumerations are read apart. */
const TOKENIZED_TYPES = new Set(['ID', 'IDREF', 'IDREFS', 'ENTITY', 'ENTITIES', 'NMTOKEN', 'NMTOKENS']);

const LIMIT_DETAILS = new Set<D>([D.XmlDepth, D.XmlTokenSize, D.XmlEntityDepth, D.XmlEntityExpansions, D.XmlEntitySize]);

export type XmlEncoding = 'utf-8' | 'utf-16le' | 'utf-16be';

/**
 * The encoding of an XML part from its first four bytes, as XML 1.0 Appendix F tells it, and the length of its byte
 * order mark. Undefined for UTF-32, which no Office application reads.
 */
export function detectEncoding(b: Uint8Array): { encoding: XmlEncoding; bom: number } | undefined {
  const at = (i: number) => (i < b.length ? b[i] : -1);
  if ((at(0) === 0 && at(1) === 0) || (at(0) === 0xff && at(1) === 0xfe && at(2) === 0 && at(3) === 0) || (at(0) === 0x3c && at(1) === 0 && at(2) === 0 && at(3) === 0)) return undefined;
  if (at(0) === 0xef && at(1) === 0xbb && at(2) === 0xbf) return { encoding: 'utf-8', bom: 3 };
  if (at(0) === 0xff && at(1) === 0xfe) return { encoding: 'utf-16le', bom: 2 };
  if (at(0) === 0xfe && at(1) === 0xff) return { encoding: 'utf-16be', bom: 2 };
  if (at(0) === 0x3c && at(1) === 0 && at(2) === 0x3f && at(3) === 0) return { encoding: 'utf-16le', bom: 0 };
  if (at(0) === 0 && at(1) === 0x3c && at(2) === 0 && at(3) === 0x3f) return { encoding: 'utf-16be', bom: 0 };
  return { encoding: 'utf-8', bom: 0 };
}

const isSpace = (c: number) => c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d;
/** A character XML 1.0 allows. The decoders leave no lone surrogates, so the halves of a pair pass here. */
const isXmlChar = (c: number) => c === 0x09 || c === 0x0a || c === 0x0d || (c >= 0x20 && c !== 0xfffe && c !== 0xffff);
const isCodePointAllowed = (c: number) => c === 0x09 || c === 0x0a || c === 0x0d || (c >= 0x20 && c <= 0xd7ff) || (c >= 0xe000 && c <= 0xfffd) || (c >= 0x10000 && c <= 0x10ffff);
/** The further normalization of XML 1.0 (3.3.3) for an attribute declared with a type other than CDATA. */
const tokenize = (v: string) => v.replace(/ {2,}/g, ' ').replace(/^ | $/g, '');

interface Open {
  el: XmlElement;
  /** Namespace bindings this element declares. */
  scope?: Map<string, string>;
}

/**
 * An internal entity has the text it expands to. An external one has the identifiers it would be fetched by, and an
 * unparsed one the notation of its data.
 */
type Entity = { text: string } | { systemId: string; publicId?: string; notation?: string };

/** What the attribute-list declarations say about one element type's attributes. */
interface ElementAttributes {
  /** Every attribute declared, since the first declaration of one binds. */
  names: Set<string>;
  /** The attributes declared with a type other than CDATA. */
  tokenized: Set<string>;
  /** The defaults, normalized, in the order they were declared. */
  defaults: Array<[string, string]>;
}

/** Where the scan of a document type declaration got to, so one that arrives in pieces is scanned once. */
interface DoctypeScan {
  /** Offset from the start of the declaration. */
  at: number;
  mode: 'outside' | 'subset' | 'declaration' | 'comment' | 'instruction';
  /** The quote of the literal being scanned, '' outside one. */
  quote: string;
}

/** What the reading of one internal subset found. */
interface Subset {
  /** A parameter entity that was not read came before, so entity and attribute-list declarations are skipped. */
  unread: boolean;
  entities: number;
  externalEntities: number;
  attributeDefaults: number;
  skipped: number;
}

/** A cursor over a whole document type declaration, or over the text of a parameter entity read inside one. */
class Cursor {
  i: number;
  /**
   * `normalized` marks the text of a parameter entity. Its line ends were normalized where it was declared, so a CR
   * in it, which only a character reference makes, stays a character of its own.
   */
  constructor(
    readonly s: string,
    start: number,
    readonly normalized: boolean,
    private readonly fail: (message: string) => never,
  ) {
    this.i = start;
  }

  /** Skips white space and tells whether there was any. */
  space(): boolean {
    const from = this.i;
    while (this.i < this.s.length && isSpace(this.s.charCodeAt(this.i))) this.i++;
    return this.i > from;
  }

  needSpace(where: string): void {
    if (!this.space()) this.fail(`No white space ${where}`);
  }

  /** Steps over `word` when it is next, and tells whether it was. */
  skip(word: string): boolean {
    if (!this.s.startsWith(word, this.i)) return false;
    this.i += word.length;
    return true;
  }

  /** Everything up to white space or a delimiter, for the caller to check as a name. */
  name(): string {
    const from = this.i;
    while (this.i < this.s.length && !NAME_END.test(this.s[this.i])) this.i++;
    return this.s.slice(from, this.i);
  }

  /** A quoted literal, without its quotes. */
  literal(what: string): string {
    const q = this.s[this.i];
    if (q !== '"' && q !== "'") this.fail(`${what} is not quoted`);
    const from = this.i + 1;
    const close = this.s.indexOf(q, from);
    if (close < 0) this.fail(`${what} is not closed`);
    this.i = close + 1;
    return this.s.slice(from, close);
  }

  /**
   * Steps past the ">" that ends an element type or notation declaration, over quoted literals. A parameter entity
   * reference may not stand inside a declaration of the internal subset, and nothing else there uses "%".
   */
  declarationEnd(): void {
    for (; this.i < this.s.length; this.i++) {
      const c = this.s[this.i];
      if (c === '>') {
        this.i++;
        return;
      }
      if (c === '%') this.fail('A parameter entity reference inside a declaration');
      if (c === '"' || c === "'") {
        const close = this.s.indexOf(c, this.i + 1);
        if (close < 0) break;
        this.i = close;
      }
    }
    this.fail('A declaration in the internal subset does not end');
  }
}

/**
 * A streaming XML 1.0 tokenizer with namespaces. It refuses references to entities nothing declares, encodings other
 * than UTF-8 and UTF-16, and anything not well-formed. It holds one markup token at a time; text goes to the handler
 * in pieces.
 *
 * A document type declaration goes to `notes.doctype`. No external DTD or other external entity is ever loaded. The
 * internal subset is read as XML 1.0 (5.1) has a processor read it that reads internal parameter entities and no
 * external ones: it keeps the general entities and attribute-list declarations, reads the text of each internal
 * parameter entity it refers to, and skips the entity and attribute-list declarations after a parameter entity it
 * does not read, unless the XML declaration says standalone="yes".
 *
 * A reference to an internal entity is expanded, within the entity limits, and the handler sees its text, markup
 * included, as if the part held it. A reference to an external entity in content goes to `notes.externalEntity`, and
 * stands in the text as UNKNOWN marks it. One in an attribute value, directly or through other entities, is refused,
 * as XML 1.0 refuses it. An attribute an element leaves out gets the default its attribute-list declaration gives, and
 * an attribute declared with a type other than CDATA is normalized further, as expat does. Each default added counts
 * as an expansion against the entity limits.
 */
export class XmlTokenizer {
  private decoder?: TextDecoder;
  private head: Uint8Array = new Uint8Array(0);
  private buf = '';
  private pos = 0;
  /** Offset in the decoded text of buf[0]. */
  private base = 0;
  private readonly stack: Open[] = [];
  private rootSeen = false;
  private doctypeSeen = false;
  /** The XML declaration says standalone="yes". */
  private standalone = false;
  private encoding: XmlEncoding = 'utf-8';
  /** Start tags read so far, which numbers the elements. */
  private elements = 0;
  /** An attribute value of the tag being read referred to an entity of the internal subset. */
  private generalReference = false;
  /** The last text piece ended in CR, so an LF that starts the next piece belongs to the same line end. */
  private lastCr = false;
  private scan?: DoctypeScan;
  /** General entities the internal subset declares, by name. */
  private readonly entities = new Map<string, Entity>();
  /** Parameter entities the internal subset declares, by name. */
  private readonly parameters = new Map<string, Entity>();
  /** What the attribute-list declarations say, by element type as written. */
  private readonly declared = new Map<string, ElementAttributes>();
  /** The entities being expanded, outermost first. A parameter entity's name has "%" in front. */
  private readonly active: string[] = [];
  /** While entity text is read, the offsets of the outermost reference, which its elements report. */
  private span?: [number, number];
  /** The stack depth at which the entity text being read started. Its end tags close nothing below it. */
  private floor = 0;
  private readonly vml: boolean;
  private readonly notes: XmlNotes;
  private readonly budget: EntityBudget;

  /**
   * `options.vml` reads the dialect Office writes in VML drawing parts, which is not well-formed XML: it skips the
   * `<![if ...]>` and `<![endif]>` markers around content and takes an unclosed `<br>` as an empty element. Both
   * branches of a marker reach the handler.
   */
  constructor(
    private readonly handler: XmlHandler,
    private readonly limits: XmlLimits,
    private readonly part: string,
    options: XmlOptions = {},
  ) {
    this.vml = options.vml ?? false;
    this.notes = options.notes ?? {};
    this.budget = options.budget ?? { expansions: 0, size: 0 };
  }

  private stop(detail: D, message: string): never {
    const category = LIMIT_DETAILS.has(detail) ? C.Limit : C.Xml;
    throw new StopError(category, detail, `${this.part}: ${message}`, { part: this.part, reason: message }, this.part);
  }

  private malformed(message: string): never {
    return this.stop(D.Malformed, message);
  }

  write(bytes: Uint8Array): void {
    if (this.decoder) {
      this.feed(this.decode(this.decoder, bytes, true), false);
      return;
    }
    this.head = this.head.length ? Buffer.concat([this.head, bytes]) : bytes;
    if (this.head.length >= 4) this.begin();
  }

  end(): void {
    const decoder = this.decoder ?? this.begin();
    this.feed(this.decode(decoder, new Uint8Array(0), false), true);
    if (this.stack.length) this.malformed(`<${this.stack[this.stack.length - 1].el.qname}> is never closed`);
    if (!this.rootSeen) this.malformed('There is no root element');
  }

  /** Picks the encoding from the first bytes, as XML 1.0 Appendix F does, and reads the bytes held so far. */
  private begin(): TextDecoder {
    const b = this.head;
    const found = detectEncoding(b);
    if (!found) return this.stop(D.Encoding, 'UTF-32 is not an encoding Office reads');
    this.encoding = found.encoding;
    const decoder = new TextDecoder(this.encoding, { fatal: true, ignoreBOM: true });
    this.decoder = decoder;
    this.head = new Uint8Array(0);
    this.feed(this.decode(decoder, b.subarray(found.bom), true), false);
    return decoder;
  }

  private decode(decoder: TextDecoder, bytes: Uint8Array, more: boolean): string {
    try {
      return decoder.decode(bytes, { stream: more });
    } catch {
      return this.stop(D.Encoding, `Bytes that are not valid ${this.encoding.toUpperCase()}`);
    }
  }

  /** Checks the XML declaration against the encoding the bytes are in, and notes whether it says standalone="yes". */
  private declaration(body: string): void {
    const version = /\bversion\s*=\s*(["'])([^"']*)\1/.exec(body)?.[2];
    if (version !== '1.0') this.malformed(`XML version ${version ?? 'missing'}`);
    this.standalone = /\bstandalone\s*=\s*(["'])yes\1/.test(body);
    const declared = /\bencoding\s*=\s*(["'])([^"']*)\1/.exec(body)?.[2]?.toLowerCase();
    if (declared === undefined) return;
    const ok = this.encoding === 'utf-8' ? declared === 'utf-8' || declared === 'utf8' : declared === 'utf-16' || declared === this.encoding;
    if (!ok) this.stop(D.Encoding, `The declaration says ${declared}, the bytes are ${this.encoding.toUpperCase()}`);
  }

  private feed(text: string, final: boolean): void {
    this.buf = this.pos < this.buf.length ? this.buf.slice(this.pos) + text : text;
    this.base += this.pos;
    this.pos = 0;
    this.loop(final);
  }

  /** Reads the text and markup in buf from pos. Entity text is read with `final` set, since all of it is there. */
  private loop(final: boolean): void {
    for (;;) {
      const lt = this.buf.indexOf('<', this.pos);
      if (lt !== this.pos) {
        // Text up to the next tag, or what can be decoded of it so far.
        let end = lt < 0 ? this.buf.length : lt;
        if (lt < 0 && !final) {
          const amp = this.buf.lastIndexOf('&');
          if (amp >= this.pos && this.buf.indexOf(';', amp) < 0) {
            if (this.buf.length - amp > MAX_REFERENCE) this.malformed('An unterminated reference');
            end = amp;
          }
        }
        if (end > this.pos) this.text(this.buf.slice(this.pos, end), this.base + this.pos);
        this.pos = end;
        if (lt < 0) break;
      }
      // At the end of the part, markup is read whole or refused, so nothing is left over.
      if (!this.markup(final)) break;
    }
  }

  /** The offset in the decoded text of buf[i], or while entity text is read, that of the reference to it. */
  private at(i: number, end = false): number {
    if (this.span) return this.span[end ? 1 : 0];
    return this.base + i;
  }

  /**
   * Character data at `offset` in the decoded text. References are expanded and line ends turned into LF. Entity
   * text had its line ends normalized where it was declared, so a CR in it, which only a character reference makes,
   * stays.
   */
  private text(raw: string, offset: number): void {
    const normalized = this.span !== undefined;
    let out = '';
    let from = 0;
    let i = 0;
    if (!normalized && this.lastCr && raw.charCodeAt(0) === 0x0a) {
      from = 1;
      i = 1;
    }
    for (; i < raw.length; i++) {
      const c = raw.charCodeAt(i);
      if (c === 0x26) {
        const semi = raw.indexOf(';', i);
        if (semi < 0 || semi - i > MAX_REFERENCE) this.malformed('An unterminated reference');
        const name = raw.slice(i + 1, semi);
        const plain = this.plainReference(name);
        if (plain !== undefined) out += raw.slice(from, i) + plain;
        else {
          this.characters(out + raw.slice(from, i));
          out = '';
          this.contentReference(name, offset + i, offset + semi + 1);
        }
        i = semi;
        from = semi + 1;
      } else if (c === 0x0d && !normalized) {
        out += `${raw.slice(from, i)}\n`;
        if (raw.charCodeAt(i + 1) === 0x0a) i++;
        from = i + 1;
      } else if (!isXmlChar(c)) {
        this.malformed(`Character U+${c.toString(16).padStart(4, '0')} is not allowed in XML`);
      }
    }
    if (!normalized) this.lastCr = raw.charCodeAt(raw.length - 1) === 0x0d;
    this.characters(from === 0 ? raw : out + raw.slice(from));
  }

  /** Hands decoded text to the handler. Outside the root element only white space may stand. */
  private characters(decoded: string): void {
    if (!this.stack.length) {
      for (let i = 0; i < decoded.length; i++) if (!isSpace(decoded.charCodeAt(i))) this.malformed('Text outside the root element');
      return;
    }
    if (decoded.length) this.handler.text?.(decoded);
  }

  /**
   * Normalizes an attribute value as XML 1.0 (3.3.3) says: references expanded, and each white space character a
   * space, except one written as a character reference. `normalized` marks entity text, whose line ends were
   * normalized where it was declared, so a CR and LF in it are two characters.
   */
  private attributeValue(raw: string, normalized: boolean): string {
    let out = '';
    let from = 0;
    for (let i = 0; i < raw.length; i++) {
      const c = raw.charCodeAt(i);
      if (c === 0x26) {
        const semi = raw.indexOf(';', i);
        if (semi < 0 || semi - i > MAX_REFERENCE) this.malformed('An unterminated reference');
        const name = raw.slice(i + 1, semi);
        out += raw.slice(from, i) + (this.plainReference(name) ?? this.attributeReference(name));
        i = semi;
        from = semi + 1;
      } else if (c === 0x0d && !normalized) {
        out += `${raw.slice(from, i)} `;
        if (raw.charCodeAt(i + 1) === 0x0a) i++;
        from = i + 1;
      } else if (c === 0x0a || c === 0x09 || c === 0x0d) {
        out += `${raw.slice(from, i)} `;
        from = i + 1;
      } else if (c === 0x3c) {
        this.malformed('"<" in an attribute value');
      } else if (!isXmlChar(c)) {
        this.malformed(`Character U+${c.toString(16).padStart(4, '0')} is not allowed in XML`);
      }
    }
    return from === 0 ? raw : out + raw.slice(from);
  }

  /** The text a predefined entity or a character reference stands for, or undefined for any other entity. */
  private plainReference(name: string): string | undefined {
    if (Object.hasOwn(PREDEFINED, name)) return PREDEFINED[name];
    if (name.startsWith('#')) return this.characterReference(name);
    if (!isNcName(name)) this.malformed('A malformed reference');
    return undefined;
  }

  private characterReference(name: string): string {
    const m = /^#(?:x([0-9a-fA-F]+)|([0-9]+))$/.exec(name);
    if (!m) this.malformed('A malformed reference');
    const code = m[1] !== undefined ? Number.parseInt(m[1], 16) : Number.parseInt(m[2], 10);
    if (!isCodePointAllowed(code)) this.malformed(`A character reference to U+${code.toString(16)}`);
    return String.fromCodePoint(code);
  }

  /** The entity a reference names. A reference to one nothing declares is refused. */
  private entity(name: string): Entity {
    const e = this.entities.get(name);
    if (!e) this.malformed(`A reference to entity "${name}", which nothing declares`);
    return e;
  }

  /** Starts expanding an entity, within the limits. */
  private enter(name: string, text: string): void {
    if (this.active.includes(name)) this.malformed(`Entity "${name}" refers to itself`);
    if (this.active.length >= this.limits.entityDepth) this.stop(D.XmlEntityDepth, `Entity references nest deeper than ${this.limits.entityDepth}`);
    this.charge(text.length);
    this.active.push(name);
  }

  /**
   * Counts one expansion of `size` characters of declared text, an entity's or an attribute default's, against the
   * limits, so a document type cannot make the part cost more than they allow.
   */
  private charge(size: number): void {
    this.budget.expansions += 1;
    if (this.budget.expansions > this.limits.entityExpansions) this.stop(D.XmlEntityExpansions, `More than ${this.limits.entityExpansions} entities and attribute defaults expanded`);
    this.budget.size += size;
    if (this.budget.size > this.limits.entitySize) this.stop(D.XmlEntitySize, `More than ${this.limits.entitySize} characters of entity text and attribute defaults expanded`);
  }

  /** A reference to a general entity in content. Its text is read as content, in place of the reference. */
  private contentReference(name: string, start: number, end: number): void {
    if (!this.stack.length) this.malformed('An entity reference outside the root element');
    const e = this.entity(name);
    const span = this.span ?? [start, end];
    if (!('text' in e)) {
      if (e.notation !== undefined) this.malformed(`A reference to entity "${name}", whose data is not XML`);
      this.notes.externalEntity?.({ name, ...(e.publicId === undefined ? {} : { publicId: e.publicId }), systemId: e.systemId, start: span[0], end: span[1] });
      this.characters(UNKNOWN + name + UNKNOWN);
      return;
    }
    this.enter(name, e.text);
    if (!this.span) this.notes.referenceStart?.(start, end);
    const { buf, pos, base, floor, span: outer } = this;
    this.buf = e.text;
    this.pos = 0;
    this.base = 0;
    this.floor = this.stack.length;
    this.span = span;
    this.loop(true);
    if (this.stack.length !== this.floor) this.malformed(`The text of entity "${name}" leaves <${this.stack[this.stack.length - 1].el.qname}> open`);
    this.buf = buf;
    this.pos = pos;
    this.base = base;
    this.floor = floor;
    this.span = outer;
    this.active.pop();
    if (!outer) this.notes.referenceEnd?.();
  }

  /** A reference to a general entity in an attribute value. Its text is normalized into the value. */
  private attributeReference(name: string): string {
    const e = this.entity(name);
    if (!('text' in e)) this.malformed(`A reference to external entity "${name}" in an attribute value`);
    this.enter(name, e.text);
    this.generalReference = true;
    const value = this.attributeValue(e.text, true);
    this.active.pop();
    return value;
  }

  /** Reads one markup token at pos. Returns false when the buffer does not yet hold all of it. */
  private markup(final: boolean): boolean {
    const b = this.buf;
    const p = this.pos;
    // Enough to tell every kind of markup apart: "<!NOTATION" is the longest opening.
    if (b.length - p < 10 && !final) return false;
    let close: number;
    if (b.startsWith('<?', p)) {
      close = b.indexOf('?>', p + 2);
      if (close < 0) return this.wait(final);
      this.instruction(b.slice(p + 2, close));
      return this.advance(close + 2);
    }
    if (b.startsWith('<!--', p)) {
      close = b.indexOf('-->', p + 4);
      if (close < 0) return this.wait(final);
      this.checkChars(b, p + 4, close);
      return this.advance(close + 3);
    }
    if (b.startsWith('<![CDATA[', p)) {
      close = b.indexOf(']]>', p + 9);
      if (close < 0) return this.wait(final);
      if (!this.stack.length) this.malformed('A CDATA section outside the root element');
      const raw = b.slice(p + 9, close);
      this.checkChars(raw, 0, raw.length);
      const t = this.span ? raw : raw.replace(/\r\n?/g, '\n');
      if (t.length) this.handler.text?.(t);
      this.lastCr = false;
      return this.advance(close + 3);
    }
    if (this.vml && b.startsWith('<![', p)) {
      close = b.indexOf(']>', p + 3);
      if (close < 0) return this.wait(final);
      const marker = b.slice(p + 3, close);
      if (!/^(if\s|endif$)/.test(marker)) this.malformed('An unknown <![ marker');
      this.checkChars(marker, 0, marker.length);
      return this.advance(close + 2);
    }
    if (b.startsWith('<!DOCTYPE', p)) {
      if (this.rootSeen || this.doctypeSeen) this.malformed('A document type declaration after the root element, or a second one');
      close = this.scanDoctype(b, p);
      if (close < 0) return this.wait(final);
      this.doctype(b.slice(p, close), this.at(p), this.at(close));
      return this.advance(close);
    }
    if (b.startsWith('<!', p)) {
      if (/^<!(ENTITY|ELEMENT|ATTLIST|NOTATION)/.test(b.slice(p, p + 10))) this.malformed('A markup declaration outside the document type declaration');
      this.malformed('An unknown <! declaration');
    }
    if (b.startsWith('</', p)) {
      close = b.indexOf('>', p + 2);
      if (close < 0) return this.wait(final);
      this.endTag(b.slice(p + 2, close), this.at(p), this.at(close + 1, true));
      return this.advance(close + 1);
    }
    // A start tag ends at the first ">" outside a quoted value.
    let i = p + 1;
    for (;;) {
      if (i >= b.length) return this.wait(final);
      const c = b.charCodeAt(i);
      if (c === 0x22 || c === 0x27) {
        const q = b.indexOf(c === 0x22 ? '"' : "'", i + 1);
        if (q < 0) return this.wait(final);
        i = q + 1;
        continue;
      }
      if (c === 0x3e) break;
      if (c === 0x3c) this.malformed('"<" inside a tag');
      i++;
    }
    this.startTag(b.slice(p + 1, i), this.at(p), this.at(i + 1, true));
    return this.advance(i + 1);
  }

  private wait(final: boolean): boolean {
    if (final) this.malformed(this.span ? 'Entity text ends inside markup' : 'The part ends inside markup');
    if (this.buf.length - this.pos > this.limits.tokenSize) this.stop(D.XmlTokenSize, `A token is longer than ${this.limits.tokenSize} characters`);
    return false;
  }

  private advance(to: number): boolean {
    if (to - this.pos > this.limits.tokenSize) this.stop(D.XmlTokenSize, `A token is longer than ${this.limits.tokenSize} characters`);
    this.pos = to;
    this.lastCr = false;
    return true;
  }

  private checkChars(s: string, from: number, to: number): void {
    for (let i = from; i < to; i++) {
      const c = s.charCodeAt(i);
      if (!isXmlChar(c)) this.malformed(`Character U+${c.toString(16).padStart(4, '0')} is not allowed in XML`);
    }
  }

  private instruction(body: string): void {
    this.checkChars(body, 0, body.length);
    const space = body.search(/\s/);
    const target = space < 0 ? body : body.slice(0, space);
    if (target.toLowerCase() === 'xml') {
      if (this.span || this.base + this.pos !== 0) this.malformed('An XML declaration that is not at the start');
      this.declaration(body);
      return;
    }
    if (!isNcName(target)) this.malformed('A processing instruction without a valid target');
  }

  /**
   * Finds the end of the document type declaration at p: the ">" outside its literals and its internal subset, where
   * comments, processing instructions and declarations may hold either. Returns the offset after it, or -1 when buf
   * does not hold all of it yet. The scan goes on from where it stopped.
   */
  private scanDoctype(b: string, p: number): number {
    const s = this.scan ?? { at: 9, mode: 'outside', quote: '' };
    this.scan = s;
    let i = p + s.at;
    for (; i < b.length; i++) {
      const c = b[i];
      if (s.quote) {
        if (c === s.quote) s.quote = '';
      } else if (s.mode === 'outside' || s.mode === 'declaration') {
        if (c === '"' || c === "'") s.quote = c;
        else if (c === '[' && s.mode === 'outside') s.mode = 'subset';
        else if (c === '>' && s.mode === 'declaration') s.mode = 'subset';
        else if (c === '>') {
          this.scan = undefined;
          return i + 1;
        }
      } else if (s.mode === 'subset') {
        if (c === ']') s.mode = 'outside';
        else if (c === '<') {
          // Nothing whole is that short, and "<!--" may not have arrived whole yet.
          if (b.length - i < 4) break;
          if (b.startsWith('<!--', i)) {
            s.mode = 'comment';
            i += 3;
          } else if (b[i + 1] === '?') {
            s.mode = 'instruction';
            i += 1;
          } else s.mode = 'declaration';
        }
      } else {
        const end = s.mode === 'comment' ? '-->' : '?>';
        if (c === end[0]) {
          if (b.length - i < end.length) break;
          if (b.startsWith(end, i)) {
            s.mode = 'subset';
            i += end.length - 1;
          }
        }
      }
    }
    s.at = i - p;
    return -1;
  }

  /** Reads a whole document type declaration, keeping the entities and attributes its internal subset declares. */
  private doctype(token: string, start: number, end: number): void {
    this.checkChars(token, 0, token.length);
    const r = new Cursor(token, 9, false, m => this.malformed(m));
    r.needSpace('after <!DOCTYPE');
    const name = r.name();
    if (!isQName(name)) this.malformed(`"${name.slice(0, 40)}" is not a valid name`);
    const id = r.space() ? this.externalId(r) : undefined;
    if (id) r.space();
    const s: Subset = { unread: false, entities: 0, externalEntities: 0, attributeDefaults: 0, skipped: 0 };
    const internalSubset = r.skip('[');
    if (internalSubset) {
      this.declarations(r, s, false);
      r.space();
    }
    // The scan ended the token at a ">", so it is the last character.
    if (r.i !== token.length - 1) this.malformed('The document type declaration does not end where it should');
    this.doctypeSeen = true;
    const { entities, externalEntities, attributeDefaults, skipped } = s;
    this.notes.doctype?.({ name, ...id, internalSubset, entities, externalEntities, attributeDefaults, skipped, start, end });
  }

  /** SYSTEM and a system literal, or PUBLIC and a public and a system literal, or undefined when neither is next. */
  private externalId(r: Cursor): { publicId?: string; systemId: string } | undefined {
    if (r.skip('SYSTEM')) {
      r.needSpace('after SYSTEM');
      return { systemId: r.literal('A system identifier') };
    }
    if (!r.skip('PUBLIC')) return undefined;
    r.needSpace('after PUBLIC');
    const publicId = r.literal('A public identifier');
    if (!PUBID.test(publicId)) this.malformed('A public identifier holds a character it may not');
    r.needSpace('after the public identifier');
    return { publicId, systemId: r.literal('A system identifier') };
  }

  /**
   * Reads markup declarations: the internal subset up to and past its "]", or with `nested` the whole text of a
   * parameter entity.
   */
  private declarations(r: Cursor, s: Subset, nested: boolean): void {
    for (;;) {
      r.space();
      if (nested && r.i === r.s.length) return;
      if (!nested && r.skip(']')) return;
      if (r.skip('<!--')) {
        const close = r.s.indexOf('-->', r.i);
        if (close < 0) this.malformed('A comment in the internal subset does not end');
        r.i = close + 3;
      } else if (r.skip('<?')) {
        const close = r.s.indexOf('?>', r.i);
        if (close < 0) this.malformed('A processing instruction in the internal subset does not end');
        const target = r.s.slice(r.i, close).split(/\s/, 1)[0];
        if (!isNcName(target) || target.toLowerCase() === 'xml') this.malformed('A processing instruction without a valid target');
        r.i = close + 2;
      } else if (r.skip('<!ENTITY')) this.entityDeclaration(r, s);
      else if (r.skip('<!ATTLIST')) this.attributeListDeclaration(r, s);
      else if (r.skip('<!ELEMENT') || r.skip('<!NOTATION')) {
        r.needSpace('after the declaration keyword');
        r.declarationEnd();
      } else if (r.skip('%')) this.parameterReference(r, s);
      else if (r.skip('<![')) this.malformed('A conditional section, which only an external subset may hold');
      else this.malformed('An unknown declaration in the internal subset');
    }
  }

  /**
   * A parameter entity reference between declarations. An internal entity's text is read as declarations, with a
   * space on each side as XML 1.0 (4.4.8) says. An external or undeclared one is not read, and the entity and
   * attribute-list declarations after it are skipped, since it could have declared the same names first.
   */
  private parameterReference(r: Cursor, s: Subset): void {
    const name = r.name();
    if (!isNcName(name) || !r.skip(';')) this.malformed('A malformed parameter entity reference');
    const e = this.parameters.get(name);
    if (!e || !('text' in e)) {
      if (!this.standalone) s.unread = true;
      return;
    }
    this.enter(`%${name}`, e.text);
    this.declarations(new Cursor(` ${e.text} `, 0, true, m => this.malformed(m)), s, true);
    this.active.pop();
  }

  /**
   * Reads an entity declaration after its keyword, and keeps the entity unless it comes after a parameter entity that
   * was not read, or gives one of the predefined names, or a name already declared: the first declaration binds.
   */
  private entityDeclaration(r: Cursor, s: Subset): void {
    r.needSpace('after <!ENTITY');
    const parameter = r.skip('%');
    if (parameter) r.needSpace('after %');
    const name = r.name();
    if (!isNcName(name)) this.malformed(`"${name.slice(0, 40)}" is not a valid entity name`);
    r.needSpace('after the entity name');
    let entity: Entity;
    const q = r.s[r.i];
    if (q === '"' || q === "'") entity = { text: this.entityValue(r.literal('An entity value'), r.normalized) };
    else {
      const id = this.externalId(r);
      if (!id) this.malformed('An entity declaration with neither a value nor an external identifier');
      entity = id;
      if (!parameter && r.space() && r.skip('NDATA')) {
        r.needSpace('after NDATA');
        const notation = r.name();
        if (!isNcName(notation)) this.malformed('NDATA without a notation name');
        entity = { ...id, notation };
      }
    }
    r.space();
    if (!r.skip('>')) this.malformed('An entity declaration does not end where it should');
    if (s.unread) {
      s.skipped++;
      return;
    }
    const map = parameter ? this.parameters : this.entities;
    if (map.has(name) || (!parameter && Object.hasOwn(PREDEFINED, name))) return;
    map.set(name, entity);
    if (parameter) return;
    s.entities++;
    if (!('text' in entity)) s.externalEntities++;
  }

  /**
   * The text an entity value expands to: line ends normalized, unless `normalized` says they were, and character
   * references replaced. References to general entities stay as written, to be expanded where the entity is used. A
   * parameter entity reference cannot stand inside a declaration in the internal subset.
   */
  private entityValue(raw: string, normalized: boolean): string {
    let out = '';
    let from = 0;
    for (let i = 0; i < raw.length; i++) {
      const c = raw.charCodeAt(i);
      if (c === 0x25) this.malformed('A parameter entity reference inside a declaration');
      else if (c === 0x26) {
        const semi = raw.indexOf(';', i);
        if (semi < 0) this.malformed('An unterminated reference');
        const name = raw.slice(i + 1, semi);
        if (name.startsWith('#')) {
          out += raw.slice(from, i) + this.characterReference(name);
          from = semi + 1;
        } else if (!isNcName(name)) this.malformed('A malformed reference');
        i = semi;
      } else if (c === 0x0d && !normalized) {
        out += `${raw.slice(from, i)}\n`;
        if (raw.charCodeAt(i + 1) === 0x0a) i++;
        from = i + 1;
      }
    }
    return out + raw.slice(from);
  }

  /**
   * Reads an attribute-list declaration after its keyword. Each attribute it names for the first time is kept, with
   * its default normalized now, as the entities declared so far give it. After a parameter entity that was not read,
   * the declaration is skipped.
   */
  private attributeListDeclaration(r: Cursor, s: Subset): void {
    r.needSpace('after <!ATTLIST');
    const element = r.name();
    if (!isQName(element)) this.malformed(`"${element.slice(0, 40)}" is not a valid name`);
    const defs: Array<[string, boolean, string | undefined]> = [];
    for (;;) {
      const spaced = r.space();
      if (r.skip('>')) break;
      if (!spaced) this.malformed('No white space before an attribute definition');
      const name = r.name();
      if (!isQName(name)) this.malformed(`"${name.slice(0, 40)}" is not a valid name`);
      r.needSpace('after the attribute name');
      const tokenized = this.attributeType(r);
      r.needSpace('after the attribute type');
      let raw: string | undefined;
      if (!r.skip('#REQUIRED') && !r.skip('#IMPLIED')) {
        if (r.skip('#FIXED')) r.needSpace('after #FIXED');
        raw = r.literal('A default value');
        if (raw.includes('<')) this.malformed('"<" in an attribute value');
      }
      defs.push([name, tokenized, raw]);
    }
    if (s.unread) {
      s.skipped++;
      return;
    }
    let declared = this.declared.get(element);
    if (!declared) {
      declared = { names: new Set(), tokenized: new Set(), defaults: [] };
      this.declared.set(element, declared);
    }
    for (const [name, tokenized, raw] of defs) {
      const value = raw === undefined ? undefined : this.attributeValue(raw, r.normalized);
      if (declared.names.has(name)) continue;
      declared.names.add(name);
      if (tokenized) declared.tokenized.add(name);
      if (value === undefined) continue;
      declared.defaults.push([name, tokenized ? tokenize(value) : value]);
      s.attributeDefaults++;
    }
  }

  /** Reads an attribute type and tells whether its values are tokens, normalized further than CDATA. */
  private attributeType(r: Cursor): boolean {
    if (r.s[r.i] === '(') {
      this.enumeration(r);
      return true;
    }
    const type = r.name();
    if (type === 'CDATA') return false;
    if (type === 'NOTATION') {
      r.needSpace('after NOTATION');
      this.enumeration(r);
    } else if (!TOKENIZED_TYPES.has(type)) this.malformed(`"${type.slice(0, 40)}" is not an attribute type`);
    return true;
  }

  /** An enumerated type: name tokens between parentheses, separated by "|". */
  private enumeration(r: Cursor): void {
    if (!r.skip('(')) this.malformed('An enumerated attribute type does not start with "("');
    for (;;) {
      r.space();
      const from = r.i;
      while (r.i < r.s.length && !/[\s|()>]/.test(r.s[r.i])) r.i++;
      if (!NMTOKEN.test(r.s.slice(from, r.i))) this.malformed('An enumerated attribute type holds something other than a name token');
      r.space();
      if (r.skip(')')) return;
      if (!r.skip('|')) this.malformed('An enumerated attribute type does not end with ")"');
    }
  }

  private lookup(prefix: string): string | undefined {
    if (prefix === 'xml') return XML_NS;
    for (let i = this.stack.length - 1; i >= 0; i--) {
      const v = this.stack[i].scope?.get(prefix);
      if (v !== undefined) return v;
    }
    return prefix === '' ? '' : undefined;
  }

  private split(qname: string): [string, string] {
    const colon = qname.indexOf(':');
    if (colon < 0) {
      if (!isNcName(qname)) this.malformed(`"${qname.slice(0, 40)}" is not a valid name`);
      return ['', qname];
    }
    const prefix = qname.slice(0, colon);
    const local = qname.slice(colon + 1);
    if (!isNcName(prefix) || !isNcName(local)) this.malformed(`"${qname.slice(0, 40)}" is not a valid name`);
    return [prefix, local];
  }

  private startTag(body: string, start: number, end: number): void {
    if (this.rootSeen && !this.stack.length) this.malformed('A second root element');
    let i = 0;
    const n = body.length;
    while (i < n && !isSpace(body.charCodeAt(i)) && body[i] !== '/') i++;
    const qname = body.slice(0, i);
    const raw: Array<[string, string]> = [];
    let selfClosing = this.vml && qname.toLowerCase() === 'br';
    this.generalReference = false;
    for (;;) {
      const before = i;
      while (i < n && isSpace(body.charCodeAt(i))) i++;
      if (i >= n) break;
      if (body[i] === '/') {
        if (i !== n - 1) this.malformed('"/" inside a tag');
        selfClosing = true;
        break;
      }
      if (i === before) this.malformed('Attributes without white space between them');
      const nameStart = i;
      while (i < n && !isSpace(body.charCodeAt(i)) && body[i] !== '=') i++;
      const name = body.slice(nameStart, i);
      while (i < n && isSpace(body.charCodeAt(i))) i++;
      if (body[i] !== '=') this.malformed(`Attribute ${name.slice(0, 40)} has no value`);
      i++;
      while (i < n && isSpace(body.charCodeAt(i))) i++;
      const quote = body[i];
      if (quote !== '"' && quote !== "'") this.malformed(`Attribute ${name.slice(0, 40)} has an unquoted value`);
      // markup() ended the tag at the first ">" outside quotes, so the value closes inside the tag.
      const close = body.indexOf(quote, i + 1);
      raw.push([name, this.attributeValue(body.slice(i + 1, close), this.span !== undefined)]);
      i = close + 1;
    }
    // Attribute-list declarations normalize the attributes they type, and add the defaults the tag leaves out. Each
    // default added counts against the entity limits, so declarations cannot multiply the work of a long part.
    const declared = this.declared.get(qname);
    let rewritten = this.generalReference;
    if (declared) {
      const given = new Set<string>();
      for (const a of raw) {
        given.add(a[0]);
        if (!declared.tokenized.has(a[0])) continue;
        const v = tokenize(a[1]);
        if (v !== a[1]) rewritten = true;
        a[1] = v;
      }
      for (const [name, value] of declared.defaults) {
        if (given.has(name)) continue;
        this.charge(value.length);
        raw.push([name, value]);
        rewritten = true;
      }
    }
    const scope = new Map<string, string>();
    const plain: Array<[string, string]> = [];
    const qnames = new Set<string>();
    for (const [name, value] of raw) {
      if (qnames.has(name)) this.malformed(`Attribute ${name.slice(0, 40)} appears twice`);
      qnames.add(name);
      if (name === 'xmlns') {
        if (value === XML_NS || value === XMLNS_NS) this.malformed('The default namespace bound to a reserved namespace');
        scope.set('', value);
      } else if (name.startsWith('xmlns:')) {
        const prefix = name.slice(6);
        if (!isNcName(prefix)) this.malformed(`"${name.slice(0, 40)}" is not a valid namespace declaration`);
        if (prefix === 'xmlns') this.malformed('The xmlns prefix declared');
        if (value === '') this.malformed(`Prefix ${prefix} bound to no namespace`);
        if ((prefix === 'xml') !== (value === XML_NS) || value === XMLNS_NS) this.malformed(`Prefix ${prefix} bound to a reserved namespace`);
        scope.set(prefix, value);
      } else plain.push([name, value]);
    }
    const depth = this.stack.length + 1;
    if (depth > this.limits.depth) this.stop(D.XmlDepth, `Elements nest deeper than ${this.limits.depth}`);
    const el: XmlElement = { ns: '', local: '', qname, attributes: [], depth, start, end, index: this.elements++, namespaces: [...scope], empty: selfClosing, rewritten };
    const open: Open = { el, scope: scope.size ? scope : undefined };
    this.stack.push(open);
    const [prefix, local] = this.split(qname);
    const ns = this.lookup(prefix);
    if (ns === undefined) this.malformed(`Prefix ${prefix} is not declared`);
    el.ns = ns;
    el.local = local;
    const expanded = new Set<string>();
    for (const [name, value] of plain) {
      const [p, l] = this.split(name);
      const ans = p === '' ? '' : this.lookup(p);
      if (ans === undefined) this.malformed(`Prefix ${p} is not declared`);
      const key = `${ans} ${l}`;
      if (expanded.has(key)) this.malformed(`Attribute ${l} appears twice in one namespace`);
      expanded.add(key);
      el.attributes.push({ ns: ans, local: l, qname: name, value });
    }
    this.rootSeen = true;
    this.handler.start?.(el);
    if (selfClosing) {
      this.stack.pop();
      this.handler.end?.(el, start, end);
    }
  }

  private endTag(body: string, start: number, end: number): void {
    let n = body.length;
    while (n > 0 && isSpace(body.charCodeAt(n - 1))) n--;
    const qname = body.slice(0, n);
    const top = this.stack.length > this.floor ? this.stack.pop() : undefined;
    if (!top) this.malformed(this.floor ? `</${qname.slice(0, 40)}> closes an element from outside the entity text it stands in` : `</${qname.slice(0, 40)}> closes nothing`);
    if (top.el.qname !== qname) this.malformed(`</${qname.slice(0, 40)}> closes <${top.el.qname.slice(0, 40)}>`);
    this.handler.end?.(top.el, start, end);
  }
}
