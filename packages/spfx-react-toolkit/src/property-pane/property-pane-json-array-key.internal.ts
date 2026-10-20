type IsUntyped<T> = 0 extends (1 & T) ? true : false;
type IsJsonArray<T, TValue> = IsUntyped<T> extends true ? false
  : [Exclude<T, undefined>] extends [never] ? false
  : [Exclude<T, undefined>] extends [readonly (infer TItem)[]]
    ? IsUntyped<TItem> extends true ? false
      : [TItem] extends [never] ? false
      : unknown extends TItem ? false
      : [TItem] extends [TValue]
        ? [TValue] extends [TItem]
          ? TValue[] extends Exclude<T, undefined> ? true : false
          : false
        : false
    : false;
/** Direct safe exact JSON-array property names shared by native fields. */
export type JsonArrayPropertyKey<TProperties extends object, TValue> = {
  [K in keyof TProperties]-?: string extends K ? never
    : K extends `${string}.${string}` | `${string}[${string}` | `${string}]${string}` | '__proto__' | 'constructor' | 'prototype' ? never
    : IsJsonArray<TProperties[K], TValue> extends true ? K : never;
}[keyof TProperties] & string;
