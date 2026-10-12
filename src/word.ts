import type { CfbEntry, CompoundFile } from './cfb';
import type { RunContext } from './context';
import { checkTime, StopError } from './errors';
import { FieldStack } from './fields';
import { Reader } from './io';
import { type StreamDecrypt, unlockDocument } from './offcrypto';
import { ByteReader, bounded, utf16 } from './oleobj';
import { type ByteSource, OfficeCategory as C, OfficeDetail as D } from './types';

/** wIdent of Word 97 and later, and of Word 6.0 and 95 (MS-DOC 2.5.2). */
const WORD8 = 0xa5ec;
const WORD6 = 0xa5dc;
/** FibRgFcLcb97 pairs (MS-DOC 2.5.6): the attached template's table and the piece table. */
const FC_STTBF_ASSOC = 32;
const FC_CLX = 33;
/** In a Word 6.0 or 95 FIB, the length of the macro table that holds WordBasic macros. Word writes 2 for an empty one. */
const WORD6_MACROS = 0x11c;
/** In a Word 6.0 or 95 FIB, fcClx and lcbClx: where the piece table of a file saved fast lies in the WordDocument stream. */
const WORD6_CLX = 0x160;
/** Text is read in pieces of at most this many characters. */
const PIECE = 65536;

const u16 = (b: Uint8Array, at: number) => b[at] | (b[at + 1] << 8);

/** A stream read through its decryption. */
const decrypted = (src: ByteSource, decrypt: StreamDecrypt): ByteSource => ({
  size: () => src.size(),
  read: async (offset, length) => decrypt(offset, await src.read(offset, length)),
});
const u32 = (b: Uint8Array, at: number) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;

/**
 * The identifier and flags of the FIB a WordDocument stream starts with (MS-DOC 2.5.2), or undefined when it does not
 * start with the FIB of Word 6.0 to 2003. The FIB's base is never encrypted.
 */
function fibBase(fib: Uint8Array): { ident: number; flags: number } | undefined {
  const ident = fib.length >= 32 ? u16(fib, 0) : 0;
  return ident === WORD8 || ident === WORD6 ? { ident, flags: u16(fib, 10) } : undefined;
}

/** Whether a WordDocument stream holds a template: its FIB has fDot set. */
export async function isWordTemplate(cfb: CompoundFile, stream: CfbEntry): Promise<boolean> {
  const base = fibBase(await cfb.source(stream).read(0, 32));
  return base !== undefined && (base.flags & 1) === 1;
}

/**
 * Inspects a Word 6.0 to 2003 document: its FIB, whether it is encrypted, the fields in its text and the template it
 * attaches. An `optional` stream that holds no FIB is not read.
 */
export async function wordRules(ctx: RunContext, cfb: CompoundFile, stream: CfbEntry, optional = false): Promise<void> {
  let source = cfb.source(stream);
  let fib = await source.read(0, 1024);
  const base = fibBase(fib);
  if (!base) {
    if (optional) return;
    throw new StopError(C.Package, D.Unrecognized, 'The WordDocument stream does not start with a Word FIB', { reason: 'a WordDocument stream without a FIB' });
  }
  const { ident, flags } = base;
  const tableEntry = cfb.child(cfb.root, flags & 0x0200 ? '1Table' : '0Table');
  let tableSource = tableEntry?.kind === 'stream' ? cfb.source(tableEntry) : undefined;
  // fEncrypted, set for RC4, RC4 CryptoAPI and XOR obfuscation alike. Word 97 to 2003 documents are read through the
  // decryption. A Word 6.0 or 95 document's encryption is another, which this reader does not undo.
  if (flags & 0x0100) {
    if (ident === WORD6 || !tableSource) throw new StopError(C.Encrypted, D.PasswordRequired, 'The document is encrypted', { stream: stream.name });
    const key = await unlockDocument(ctx, fib, tableSource, stream.name);
    source = decrypted(source, key.word);
    tableSource = decrypted(tableSource, key.table);
    fib = await source.read(0, 1024);
  }
  const location = stream.name;
  const fields = new FieldStack(ctx, location);
  // Word marks a field's start, separator and end with 0x13, 0x14 and 0x15 in the text. Other control characters in
  // a field's code, such as the 0x01 of a hyperlink's data, are not part of what it says.
  const text = (s: string) => {
    for (const ch of s) {
      if (ch === '\x13') fields.begin();
      else if (ch === '\x14') fields.separate();
      else if (ch === '\x15') fields.end();
      else if (ch >= ' ' || ch === '\t') {
        if (fields.inCode) fields.text(ch);
        else fields.result(ch);
      }
    }
  };
  const reader = new Reader(source, stream.size, 65536, 8);
  if (ident === WORD6) {
    // Word 6.0 and 95 keep their text as 8-bit characters. A file saved fast (fComplex) adds pieces to the end, and its
    // piece table, in the WordDocument stream itself at fcClx (FIB offset 0x160), gives their order. Any other file,
    // or one whose piece table does not read, is read from fcMin to fcMac, every byte between, as msodde does.
    const clxAt = u32(fib, WORD6_CLX);
    const clxSize = u32(fib, WORD6_CLX + 4);
    const list = flags & 0x0004 && clxSize > 0 && clxAt + clxSize <= stream.size ? pieces(await reader.read(clxAt, clxSize)).filter(p => !p.compressed && p.fc + p.length <= stream.size) : [];
    if (list.length) {
      for (const p of list) {
        for (let cp = 0; cp < p.length; cp += PIECE) {
          checkTime(ctx.settings.deadline);
          text(latin(await reader.read(p.fc + cp, Math.min(PIECE, p.length - cp))));
        }
      }
    } else {
      const from = u32(fib, 0x18);
      const to = Math.min(u32(fib, 0x1c), stream.size);
      for (let at = from; at < to; at += PIECE) {
        checkTime(ctx.settings.deadline);
        text(latin(await reader.read(at, Math.min(PIECE, to - at))));
      }
    }
    if (fib.length >= WORD6_MACROS + 4 && u32(fib, WORD6_MACROS) > 2) ctx.report(C.Macro, D.WordBasic, location, { stream: location, bytes: u32(fib, WORD6_MACROS) });
  } else {
    const pairs = fcLcb(fib);
    const table = tableEntry;
    if (table?.kind === 'stream' && tableSource) {
      const clx = pairs[FC_CLX];
      if (clx && clx.lcb > 0 && clx.fc + clx.lcb <= table.size) {
        for (const p of pieces(await tableSource.read(clx.fc, clx.lcb))) {
          for (let cp = 0; cp < p.length; cp += PIECE) {
            checkTime(ctx.settings.deadline);
            const n = Math.min(PIECE, p.length - cp);
            if (p.compressed) text(latin(await reader.read(p.fc + cp, n)));
            else text(utf16(await reader.read(p.fc + cp * 2, n * 2)));
          }
        }
      }
      const assoc = pairs[FC_STTBF_ASSOC];
      if (assoc && assoc.lcb > 0 && assoc.fc + assoc.lcb <= table.size) {
        // The attached template is string 1 of SttbfAssoc (MS-DOC 2.9.271).
        const template = sttb(await tableSource.read(assoc.fc, Math.min(assoc.lcb, 65536)))[1] ?? '';
        if (template.trim()) ctx.external(D.Template, template, location, { source: 'SttbfAssoc' }, 'template');
      }
    }
  }
  fields.finish();
}

/** The fc and lcb pairs of a Word 97 FIB, after its variable-length arrays of 16- and 32-bit values. */
function fcLcb(fib: Uint8Array): Array<{ fc: number; lcb: number }> {
  const r = new ByteReader(fib);
  const out: Array<{ fc: number; lcb: number }> = [];
  return bounded(
    () => {
      r.skip(32);
      r.skip(r.u16() * 2);
      r.skip(r.u16() * 4);
      const n = r.u16();
      for (let i = 0; i < n; i++) out.push({ fc: r.u32(), lcb: r.u32() });
      return out;
    },
    () => out,
  );
}

/** The pieces of a Clx (MS-DOC 2.9.38): where each run of text is, and whether it is stored 8 bits a character. */
function pieces(clx: Uint8Array): Array<{ fc: number; length: number; compressed: boolean }> {
  const r = new ByteReader(clx);
  const out: Array<{ fc: number; length: number; compressed: boolean }> = [];
  return bounded(
    () => {
      // Prc entries of formatting come first, each with its size.
      while (r.left > 0 && clx[r.pos] === 0x01) {
        r.skip(1);
        r.skip(r.u16());
      }
      if (r.u8() !== 0x02) return out;
      const lcb = r.u32();
      const n = Math.floor((lcb - 4) / 12);
      const plc = new ByteReader(r.bytes(Math.min(lcb, r.left)));
      const cps: number[] = [];
      for (let i = 0; i <= n; i++) cps.push(plc.u32());
      for (let i = 0; i < n; i++) {
        plc.skip(2);
        const fc = plc.u32();
        plc.skip(2);
        const length = cps[i + 1] - cps[i];
        if (length <= 0) continue;
        const compressed = (fc & 0x40000000) !== 0;
        out.push({ fc: compressed ? (fc & 0x3fffffff) / 2 : fc, length, compressed });
      }
      return out;
    },
    () => [],
  );
}

/** The strings of an STTB (MS-DOC 2.2.4), in 16-bit characters when it starts with 0xFFFF, as Word 97 writes one. */
function sttb(b: Uint8Array): string[] {
  const r = new ByteReader(b);
  const out: string[] = [];
  return bounded(
    () => {
      const wide = r.u16() === 0xffff;
      if (!wide) r.pos = 0;
      const count = r.u16();
      const extra = r.u16();
      for (let i = 0; i < count; i++) {
        const n = wide ? r.u16() : r.u8();
        out.push(wide ? utf16(r.bytes(n * 2)) : latin(r.bytes(n)));
        r.skip(extra);
      }
      return out;
    },
    () => out,
  );
}

const latin = (b: Uint8Array) => Buffer.from(b.buffer, b.byteOffset, b.byteLength).toString('latin1');
