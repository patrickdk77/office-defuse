/** BrtBeginSupBook, the record that opens each link of an .xlsb external link part. */
const BEGIN_SUPBOOK = 360;
/** The link type BrtBeginSupBook gives a DDE link. */
const SUPBOOK_DDE = 1;
/** How much of one BrtBeginSupBook is kept: far more than the two names of a DDE link need. */
const KEEP = 65536;

export interface DdeLink {
  application: string;
  topic: string;
}

/**
 * The record header at the start of `b`: the record's type and size, undefined while `b` is too short, or 'bad' when
 * the header runs longer than the format allows. A type takes one or two bytes and a size one to four, seven bits to
 * a byte, and a byte with its high bit clear ends each.
 */
function header(b: number[]): { type: number; size: number } | 'bad' | undefined {
  let i = 0;
  let type = 0;
  for (let shift = 0; ; shift += 7) {
    if (i >= b.length) return undefined;
    type |= (b[i] & 0x7f) << shift;
    if (!(b[i++] & 0x80)) break;
    if (shift === 7) return 'bad';
  }
  let size = 0;
  for (let shift = 0; ; shift += 7) {
    if (i >= b.length) return undefined;
    size += (b[i] & 0x7f) * 2 ** shift;
    if (!(b[i++] & 0x80)) break;
    if (shift === 21) return 'bad';
  }
  return { type, size };
}

/** An XLWideString at `at`: a 32-bit character count, then UTF-16LE. Cut short at the end of what was kept. */
function wideString(b: Buffer, at: number): { text: string; next: number } {
  if (at + 4 > b.length) return { text: '', next: b.length };
  const end = Math.min(b.length, at + 4 + b.readUInt32LE(at) * 2);
  return { text: b.toString('utf16le', at + 4, end), next: end };
}

/**
 * Reads the records of an .xlsb external link part as its chunks arrive, and collects its DDE links. Each record
 * body is skipped, except the first 64 KiB of a BrtBeginSupBook. A header that breaks the record format ends the
 * reading; Excel cannot read such a part either.
 */
export class XlsbLinks {
  readonly dde: DdeLink[] = [];
  private head: number[] = [];
  private left = 0;
  private body: Buffer[] | undefined;
  private kept = 0;
  private broken = false;

  write(chunk: Uint8Array): void {
    let i = 0;
    while (i < chunk.length && !this.broken) {
      if (this.left > 0) {
        const n = Math.min(this.left, chunk.length - i);
        if (this.body && this.kept < KEEP) {
          const part = Buffer.from(chunk.subarray(i, i + Math.min(n, KEEP - this.kept)));
          this.body.push(part);
          this.kept += part.length;
        }
        this.left -= n;
        i += n;
        if (this.left === 0) this.finish();
        continue;
      }
      this.head.push(chunk[i++]);
      const h = header(this.head);
      if (h === undefined) continue;
      if (h === 'bad') {
        this.broken = true;
        return;
      }
      this.head = [];
      this.body = h.type === BEGIN_SUPBOOK ? [] : undefined;
      this.kept = 0;
      this.left = h.size;
      if (this.left === 0) this.finish();
    }
  }

  private finish(): void {
    if (!this.body) return;
    const b = Buffer.concat(this.body);
    this.body = undefined;
    if (b.length < 2 || b.readUInt16LE(0) !== SUPBOOK_DDE) return;
    const application = wideString(b, 2);
    this.dde.push({ application: application.text, topic: wideString(b, application.next).text });
  }
}
