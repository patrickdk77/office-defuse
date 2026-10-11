import { crc32 } from './crc32';

/** The dictionary MS-OXRTFCP 2.1.3.1.2 starts from: 207 bytes of RTF that most bodies begin with. */
const PRESET = Buffer.from(
  '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx',
  'latin1',
);
/** COMPTYPE: LZFu for compressed, MELA for stored. */
const COMPRESSED = 0x75465a4c;
const UNCOMPRESSED = 0x414c454d;

/**
 * The CRC of MS-OXRTFCP 2.1.3.2: CRC-32's table, started at 0 and never inverted. zlib's CRC-32 inverts the value it
 * starts from and the value it ends with, so starting it from all ones and inverting its result gives this one.
 */
const rtfCrc = (b: Uint8Array) => ~crc32(b, 0xffffffff) >>> 0;

/**
 * Decompresses an Outlook message's PidTagRtfCompressed (MS-OXRTFCP 2.1.3.1). Returns undefined for input the
 * specification calls corrupt: cut short, of an unknown type, with a CRC that does not match, or longer than `limit`.
 */
export function decompressRtf(b: Uint8Array, limit: number): Uint8Array | undefined {
  if (b.length < 16) return undefined;
  const view = Buffer.from(b.buffer, b.byteOffset, b.byteLength);
  const compSize = view.readUInt32LE(0);
  const rawSize = view.readUInt32LE(4);
  const type = view.readUInt32LE(8);
  const crc = view.readUInt32LE(12);
  if (compSize < 12 || compSize + 4 > b.length || rawSize > limit) return undefined;
  const contents = view.subarray(16, 4 + compSize);
  if (type === UNCOMPRESSED) return contents.subarray(0, rawSize);
  if (type !== COMPRESSED || rtfCrc(contents) !== crc) return undefined;
  const dict = Buffer.alloc(4096);
  PRESET.copy(dict);
  let write = PRESET.length;
  const out = Buffer.alloc(rawSize);
  let n = 0;
  let at = 0;
  while (at < contents.length) {
    const control = contents[at++];
    for (let bit = 0; bit < 8; bit++) {
      if ((control & (1 << bit)) === 0) {
        if (at >= contents.length) return undefined;
        const c = contents[at++];
        if (n < rawSize) out[n++] = c;
        dict[write] = c;
        write = (write + 1) & 4095;
        continue;
      }
      if (at + 2 > contents.length) return undefined;
      const ref = (contents[at] << 8) | contents[at + 1];
      at += 2;
      let read = ref >> 4;
      if (read === write) return out.subarray(0, n);
      for (let k = (ref & 0x0f) + 2; k > 0; k--) {
        const c = dict[read];
        read = (read + 1) & 4095;
        if (n < rawSize) out[n++] = c;
        dict[write] = c;
        write = (write + 1) & 4095;
      }
    }
  }
  return undefined;
}

/** RTF stored as a PidTagRtfCompressed value of type UNCOMPRESSED, whose CRC is 0 (MS-OXRTFCP 2.1.3.1.1). */
export function storeRtf(rtf: Uint8Array): Uint8Array {
  const head = Buffer.alloc(16);
  head.writeUInt32LE(rtf.length + 12, 0);
  head.writeUInt32LE(rtf.length, 4);
  head.writeUInt32LE(UNCOMPRESSED, 8);
  return Buffer.concat([head, rtf]);
}
