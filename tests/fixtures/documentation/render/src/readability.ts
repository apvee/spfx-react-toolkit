/** Ownership readability fixture.
 * @packageDocumentation
 * @docPage readability
 */
/** Canonical generic owner. */
export interface ExampleBase<T> {
  /** Optional source value. Omitted means no value.
   * @defaultValue undefined
   * @remarks The source owner keeps this explanation.
   * @example Source value
   * ```ts
   * const sourceValue = 'owned';
   * ```
   */
  readonly value?: T;
  /** Read a value.
   * @param input - The input value.
   * @returns The observed value.
   * @example Source reader
   * ```ts
   * const sourceReader = 'owned';
   * ```
   */
  read(input: T): T;
}
/** Instantiated inherited view. */
export interface ExampleChild extends ExampleBase<string> {}
/** Mapped mutable view. */
export type ExampleMutable = { -readonly [K in keyof ExampleBase<number>]-?: ExampleBase<number>[K] };
/** Equal examples on an independent source owner. */
export interface ExampleIndependent {
  /** Independently authored source value.
   * @example Source value
   * ```ts
   * const sourceValue = 'owned';
   * ```
   */
  value: string;
}
/** Local override must stay local. */
export interface ExampleOverride extends ExampleBase<string> {
  /** Local value. Omitted means a local value.
   * @defaultValue local
   * @remarks Local explanation.
   * @example Local value
   * ```ts
   * const localValue = 'local';
   * ```
   */
  value?: string;
}
/** Anonymous returns and branches keep their complete inventories. */
export declare function exampleFactory(): ExampleBase<string> & { forbidden?: never };
/** Union alternatives remain separate. */
export type ExampleBranches = { mode: 'disabled'; value?: never } | { mode: 'enabled'; value: string };
