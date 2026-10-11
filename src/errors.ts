import type { OfficeCategory, OfficeDetail } from './types';

/** The run passed its deadline. The code makes every package of the family pass it up instead of treating it as one plugin failing. */
export class TimeLimitError extends Error {
  readonly code = 'DEFUSE_LIMIT';
  readonly limit = 'time';
}

/** A file sits deeper in the upload than `limits.nestingDepth` allows. The code makes every package pass it up. */
export class NestingLimitError extends Error {
  readonly code = 'DEFUSE_LIMIT';
  readonly limit = 'nesting';
}

/** An I/O error inside a nested file's own run, carried up so the upload's run fails on it as on its own. */
export class NestedIoError extends Error {
  readonly code = 'DEFUSE_IO';
  constructor(cause: unknown) {
    super(String((cause as Error)?.message ?? cause), { cause });
  }
}

/**
 * The code of an error that stops the whole tree of files, from any package: DEFUSE_LIMIT for a time or nesting
 * limit, DEFUSE_IO for an I/O error. Packages cannot share error classes, so the code alone marks one.
 */
export function stopsTree(e: unknown): 'DEFUSE_LIMIT' | 'DEFUSE_IO' | undefined {
  const code = typeof e === 'object' && e !== null ? (e as { code?: unknown }).code : undefined;
  return code === 'DEFUSE_LIMIT' || code === 'DEFUSE_IO' ? code : undefined;
}

/** An error from the operating system, such as EIO or ENOSPC, which says nothing about the file. */
export const isIoError = (e: unknown) => typeof (e as { code?: unknown })?.code === 'string' && /^E[A-Z]+$/.test((e as { code: string }).code);

/**
 * A refusal or a limit that stops the run. The engine turns it into the finding it names and rejects the file with
 * no score, whatever the overrides say.
 */
export class StopError extends Error {
  constructor(
    readonly category: OfficeCategory,
    readonly detail: OfficeDetail,
    message: string,
    readonly data?: Record<string, string | number>,
    readonly location?: string,
  ) {
    super(message);
  }
}

/** Throws the time limit once `deadline`, a Date.now() value, has passed. */
export function checkTime(deadline: number | undefined): void {
  if (deadline !== undefined && Date.now() > deadline) throw new TimeLimitError('Time limit exceeded');
}
