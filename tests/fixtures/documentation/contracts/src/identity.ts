/**
 * Preserves a typed value.
 * @useWhen You need a generic identity callback.
 * @template T - Type of the supplied value.
 * @param value - Value to return.
 * @returns The supplied value.
 * @example Identity
 * ```ts
 * import { identity } from '@fixture/contracts';
 * identity(1);
 * ```
 */
export const identity = <T>(value: T): T => value;
