/**
 * Returns the supplied count; see {@link count}.
 * @docPage helpers/count
 * @useWhen You already own the count.
 * @param value - The count to return.
 * @returns The supplied count.
 * @example Count
 * ```ts
 * import { count } from '@fixture/artifacts';
 * const result: number = count(3);
 * ```
 */
export function count(value: number): number { return value; }
