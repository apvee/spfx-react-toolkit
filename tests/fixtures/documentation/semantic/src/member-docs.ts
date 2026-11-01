/** Public owner; must never become an undocumented callback's summary. */
export interface CallbackOwner {
  /** Pure synchronous source returning original values. */
  accessor: (item: number) => string;
  /** Tooltip for the original value. */
  tooltip?: string | ((item: number) => string);
  /** Outer property contract. */
  explicit: /** Explicit anonymous signature contract. */ (item: number) => string;
  /** Nested configuration contract. */
  config: { readonly label: string };
  undocumented: (item: number) => string;
  /** Outer parenthesized property. */
  parenthesized: /** Explicit parenthesized type. */ ((item: number) => string);
  /** Outer union property. */
  union: /** Explicit union type. */ (string | ((item: number) => string));
  /** Outer intersection property. */
  intersection: /** Explicit intersection type. */ (((item: number) => string) & { readonly tag: string });
  /** Outer nested property. */
  nested: /** Explicit outer wrapper type. */ (
    /** Explicit nearest wrapper type. */ ((item: number) => string));
}
