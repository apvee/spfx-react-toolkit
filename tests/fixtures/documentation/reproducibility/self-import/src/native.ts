/**
 * Options from the dedicated native entry.
 * @docPage helpers/native
 * @example Options
 * ```ts
 * import type { NativeOptions } from '@fixture/cold/native';
 * const options: NativeOptions = { name: 'cold' };
 * ```
 */
export interface NativeOptions {
  /** Required name. */
  name: string;
}
