import type { RunState } from './context';
import { checkTime } from './errors';
import type { RtfCut } from './plan';
import type { ByteSink, ByteSource } from './types';

/** Sorts cuts and drops each one that lies inside another, so they apply in one pass. */
function normalize(cuts: RtfCut[]): RtfCut[] {
  const sorted = [...cuts].sort((a, b) => a.start - b.start || b.end - a.end);
  const out: RtfCut[] = [];
  let reach = 0;
  for (const c of sorted) {
    if (c.start < reach) continue;
    out.push(c);
    reach = Math.max(reach, c.end);
  }
  return out;
}

/** Copies the bytes from `from` to `to` of the source. */
async function copyRange(run: RunState, source: ByteSource, from: number, to: number, sink: ByteSink): Promise<void> {
  for (let at = from; at < to; at += 262144) {
    checkTime(run.deadline);
    const want = Math.min(262144, to - at);
    const b = await source.read(at, want);
    if (b.length !== want) throw Object.assign(new Error(`The source returned ${b.length} bytes at offset ${at} where ${want} were expected`), { code: 'EIO' });
    await sink.write(b);
  }
}

/**
 * Writes an RTF file without the ranges the cuts name: the groups of the destinations a strip finding names, and any
 * bytes after the group that closes the document. A cut can put text in place of what it removes, such as an object's
 * result group in place of the object. Everything else keeps its bytes.
 */
export async function writeRtf(run: RunState, source: ByteSource, size: number, cuts: RtfCut[], sink: ByteSink): Promise<void> {
  let pos = 0;
  for (const c of normalize(cuts)) {
    await copyRange(run, source, pos, Math.min(c.start, size), sink);
    if (c.text) await sink.write(c.text);
    if (c.copy) await copyRange(run, source, c.copy[0], Math.min(c.copy[1], size), sink);
    pos = Math.max(pos, c.end);
  }
  await copyRange(run, source, pos, size, sink);
}
