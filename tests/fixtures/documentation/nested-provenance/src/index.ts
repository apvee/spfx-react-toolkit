/** Nested configuration.
 */
export interface Configuration {
  /** Cache settings. */
  cache?: {
    /** Cache store. Omit for session storage.
     * @defaultValue 'session'
     * @remarks
     * ### Requirements
     * Keep this nested cache setting local.
     * @example Cache store
     * ```ts
     * import type { Configuration } from '@fixture/nested-provenance';
     * const configuration: Configuration = { cache: { storage: 'local' } };
     * ```
     */
    storage?: 'session' | 'local';
    /** Key builder. Omit for the native key. */
    keyFactory?: (url: string) => string;
  };
}
/** Token provider.
 */
export interface TokenProvider {
  /** Requests a token.
   * @param options - Omit for native options.
   * @returns The token.
   */
  getToken(options?: {
    /** Claims challenge. Omit for no challenge. */
    readonly claims?: string;
    /** Cached token preference. Omit for native behavior.
     * @defaultValue true
     */
    readonly useCachedToken?: boolean;
  }): Promise<string>;
}
/** Same-line parent and nested contracts.
 */
export interface Collision {
  /** Parent mode. Omit for parent. @defaultValue 'parent-mode' @remarks Parent mode requirements. */ mode?: string; nested: { /** Nested mode. Omit for nested. @defaultValue 'nested-mode' @remarks Nested mode requirements. */ mode?: string; plain?: string };
  /** Parent state. Omit for parent state.
   * @defaultValue 'parent'
   * @remarks
   * ### Requirements
   * Parent state requirements.
   */
  state?: string;
  /** Nested state settings. */
  options: {
    /** Nested state. Omit for nested state.
     * @defaultValue 'nested'
     * @remarks
     * ### Requirements
     * Nested state requirements.
     * @example Nested state
     * ```ts
     * import type { Collision } from '@fixture/nested-provenance';
     * const collision: Collision = { options: { state: 'nested' }, nested: {} };
     * ```
     */
    state?: string;
  };
}
/** Generic base.
 * @typeParam Value - The stored value.
 */
export interface Base<Value> {
  /** Stored value. Omit for the base value.
   * @defaultValue The base value.
   * @remarks
   * ### Requirements
   * Base value requirements.
   */
  value?: Value;
}
/** Genuine inheritance.
 */
export interface Derived extends Base<string> {}
/** Local override.
 */
export interface Override extends Base<string> {
  /** Local value. Omit for local.
   * @defaultValue 'local'
   * @remarks
   * ### Requirements
   * Local value requirements.
   */
  value?: string;
}
/** Genuine mapped ownership.
 */
export type Mapped = Readonly<Pick<Base<number>, 'value'>>;
/** Returns a real named interface.
 * @returns A named base value.
 */
export declare function getBase(): Base<boolean>;
