import type { OfficeCategory, OfficeDetail } from './types';

/** The run passed its deadline. The code makes every package of the family pass it up instead of treating it as one plugin failing. */
export class TimeLimitError extends Error {
  readonly code = 'DEFUSE_LIMIT';
  readonly limit = 'time';
}

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
