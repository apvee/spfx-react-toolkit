/** Selects a required scoped value.
 * @useWhen Select a value directly.
 * @typeParam TSelected - Required selected value type.
 * @param options - Required selected value.
 * @returns The selected value.
 * @example Scoped
 * ```ts
 * import { select } from '@fixture/overload-contracts';
 * const value: number = select(42);
 * ```
 */
export declare function select<TSelected>(options: TSelected): TSelected;
/** Selects a legacy optional value for a site.
 * @typeParam TLegacy - Optional legacy value type.
 * @param siteUrl - Optional site URL; omitted or undefined leaves selection idle.
 * @param options - Optional legacy value; omitted or undefined returns undefined.
 * @returns The legacy value, or undefined when omitted.
 * @example Legacy
 * ```ts
 * import { select } from '@fixture/overload-contracts';
 * const value: boolean | undefined = select('https://example.test', true);
 * ```
 */
export declare function select<TLegacy>(siteUrl?: string, options?: TLegacy): TLegacy | undefined;
/** Selects through a local alias with both call forms.
 * @useWhen Use the alternate spelling.
 * @param siteUrl - Legacy site URL; omitted or undefined leaves selection idle.
 * @param options - Scoped or legacy value; omitted or undefined in the legacy form returns undefined.
 * @returns The selected value.
 * @example
 * ```ts
 * import { alternate } from '@fixture/overload-contracts';
 * const value: number = alternate(42);
 * ```
 */
export { select as alternate };
