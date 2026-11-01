/** Source measure namespace purpose.
 * @packageDocumentation
 * @docPage shared
 * @useWhen Construct measure values.
 * @remarks Namespace requirements belong here only.
 * @example
 * ```ts
 * import { measures } from '@fixture/pilot';
 * const value = measures.unit(1);
 * ```
 */
/** Returns its measure.
 * @param value - Required measure.
 * @returns The measure.
 */
export function unit(value: number): number { return value; }
/** Declared nested namespace owns its purpose. */
export namespace Nested {
  /** Nested value. */
  export const value = 1;
}
/** Declared module-named namespace purpose. */
export namespace module {
  /** Declared namespace value. */
  export const value = 2;
}
