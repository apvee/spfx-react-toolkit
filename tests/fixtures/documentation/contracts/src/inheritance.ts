/** Original options.
 * @example Original
 * ```ts
 * import type { Original } from '@fixture/contracts';
 * const original: Original = {};
 * ```
 */
export interface Original {
  /** Upstream label. If omitted, no label is shown.
   * @defaultValue No label when omitted.
   */
  label?: string;
}
/** Derived options.
 * @example Derived
 * ```ts
 * import type { Derived } from '@fixture/contracts';
 * const derived: Derived = {};
 * ```
 */
export interface Derived extends Original {
  /** Local size. When omitted, size follows the host.
   * @remarks
   * ### Inherited behavior
   * Uses Original's label contract.
   */
  size?: number;
}
/** Custom label options.
 * @example Customized
 * ```ts
 * import type { Customized } from '@fixture/contracts';
 * const customized: Customized = {};
 * ```
 */
export interface Customized extends Original {
  /** Local label. When omitted, a local label is used.
   * @defaultValue Local label.
   * @remarks
   * ### Requirements
   * Requires the local label context.
   */
  label?: string;
}
