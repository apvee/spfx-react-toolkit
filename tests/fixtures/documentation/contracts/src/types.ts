/**
 * Navigation types.
 * @packageDocumentation
 * @docPage controls/pagination
 */
/**
 * Input choices for paging.
 * @example Options
 * ```ts
 * import type { PagingOptions } from '@fixture/contracts';
 * const options: PagingOptions = {};
 * ```
 */
export interface PagingOptions {
  /** Enables navigation. When omitted or undefined, navigation is enabled.
   * @default true
   */
  enabled?: boolean;
}
