import type { CfbEntry, CompoundFile } from './cfb';
import { type CfbNode, type CfbTree, writeCfb } from './cfb-writer';
import type { RunState } from './context';
import { CompoundPlan } from './plan';
import type { ByteSink, ByteSource } from './types';

/** The nodes of a storage's subtree, without what the plan drops, with what plugins returned in place of the rest. */
function nodes(cfb: CompoundFile, storage: CfbEntry, plan: CompoundPlan, opened: ByteSource[]): CfbNode[] {
  const out: CfbNode[] = [];
  for (const c of storage.children) {
    if (plan.drop.has(c)) continue;
    const replaced = plan.replace.get(c);
    if (c.kind === 'stream') {
      if (!replaced) {
        out.push({ kind: 'stream', name: c.name, size: c.size, source: cfb.source(c) });
        continue;
      }
      const source = replaced.source();
      opened.push(source);
      out.push({ kind: 'stream', name: c.name, size: replaced.length, source });
      continue;
    }
    const graft = plan.grafts.get(c);
    if (graft)
      out.push({ kind: 'storage', name: c.name, clsid: graft.root.clsid === '00000000-0000-0000-0000-000000000000' ? '' : graft.root.clsid, children: nodes(graft, graft.root, plan, opened) });
    else out.push({ kind: 'storage', name: c.name, clsid: c.clsid === '00000000-0000-0000-0000-000000000000' ? '' : c.clsid, children: nodes(cfb, c, plan, opened) });
  }
  return out;
}

/**
 * Writes a compound file without the streams and storages the plan drops. A stream a plugin replaced takes the bytes
 * the plugin wrote, and an object's storage a plugin replaced takes the storages and streams of the compound file the
 * plugin wrote. Every other stream is copied as it is. Sectors no stream holds, the rest of each stream's last sector
 * and anything after the last sector are not written.
 */
export async function writeCompound(run: RunState, cfb: CompoundFile, plan: CompoundPlan, sink: ByteSink): Promise<void> {
  const clsid = cfb.root.clsid === '00000000-0000-0000-0000-000000000000' ? '' : cfb.root.clsid;
  const opened: ByteSource[] = [];
  try {
    await writeCfb({ clsid, children: nodes(cfb, cfb.root, plan, opened) }, sink, run.deadline);
  } finally {
    for (const s of opened) await s.close?.();
  }
}

/** A storage and everything under it, as the root of a compound file of its own. */
export function subtree(cfb: CompoundFile, storage: CfbEntry): CfbTree {
  return { clsid: storage.clsid === '00000000-0000-0000-0000-000000000000' ? '' : storage.clsid, children: nodes(cfb, storage, new CompoundPlan(), []) };
}
