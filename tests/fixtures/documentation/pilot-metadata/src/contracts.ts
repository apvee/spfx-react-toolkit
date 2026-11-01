/** Callable construction type. See {@link compose} and {@link Compose}.
 * @docPage shared
 * @example
 * ```ts
 * import { compose, Compose } from '@fixture/pilot';
 * const factory: Compose = compose;
 * ```
 */
export interface Compose {
  /** Creates an empty result.
   * @returns Empty result.
   */
  (): string;
  /** Preserves a supplied value.
   * @typeParam TValue - Required value type.
   * @param value - Required original value.
   * @returns Original value.
   */
  <TValue>(value: TValue): TValue;
  /** Configures a result.
   * @typeParam TItem - Result item type.
   * @param options - Exactly one required result configuration.
   * @returns Configured result.
   */
  <TItem>(options: {
    /** Required item. */
    item: TItem;
  } | {
    /** Required items. */
    items: TItem[];
  }): TItem;
}
/** Preserves both alternatives of the named choice.
 * @typeParam T - Choice value type.
 * @example
 * ```ts
 * const example = { kind: 'one', value: 1 };
 * ```
 */
type Choice<T> = {
  /** First discriminant. */
  kind: 'one';
  /** Required value. */
  value: T;
} | {
  /** Second discriminant. */
  kind: 'many';
  /** Required values. */
  values: T[];
};
/** Resolves a choice.
 * @docPage shared
 * @useWhen Read either alternative.
 * @typeParam T - Choice value type.
 * @param choice - Required choice.
 * @returns Chosen value.
 * @example
 * ```ts
 * import { choose } from '@fixture/pilot';
 * const value = choose({ kind: 'one', value: 1 });
 * ```
 */
export declare function choose<T>(choice: Choice<T>): T;
/** Retains outer and inner inference.
 * @docPage shared
 * @typeParam TOuter - Required outer model.
 * @typeParam TInner - Required selected key.
 * @param value - Required selected key.
 * @returns Selected key.
 * @example
 * ```ts
 * import { Selector } from '@fixture/pilot';
 * const select: Selector<{ name: string }> = value => value;
 * ```
 */
export type Selector<TOuter> = <TInner extends keyof TOuter>(value: TInner) => TInner;
/** Documents only the owning outer generic here.
 * @docPage shared
 * @typeParam TOuter - Required outer model.
 * @example
 * ```ts
 * import { SeparateSelector } from '@fixture/pilot';
 * const select: SeparateSelector<{ name: string }> = value => value;
 * ```
 */
export interface SeparateSelector<TOuter> {
  /** Owns the inner call contract.
   * @typeParam TInner - Required selected key.
   * @param value - Required selected key.
   * @returns Selected key.
   */
  <TInner extends keyof TOuter>(value: TInner): TInner;
}
/** Adds a field while preserving the hidden choice contract.
 * @docPage shared
 * @typeParam T - Choice value type.
 * @example
 * ```ts
 * import { ExtendedChoice } from '@fixture/pilot';
 * const value: ExtendedChoice<number> = { kind: 'one', value: 1, extra: true };
 * ```
 */
export type ExtendedChoice<T> = Choice<T> & {
  /** Extra field. */
  extra: boolean;
};
/** Configures an explicitly documented shape.
 * @docPage shared
 * @useWhen Configure the nested value.
 * @param options - Required parent parameter description.
 * @returns Nothing.
 * @example
 * ```ts
 * import { configure } from '@fixture/pilot';
 * configure({ value: 1 });
 * ```
 */
export declare function configure(options: /** Explicit parameter type purpose. */ {
  /** Required value. */
  value: number;
}): void;

/** Describes a reusable model handler.
 * @docPage shared
 * @typeParam TOuter - Required enclosing handler model.
 * @example
 * ```ts
 * import { ModelHandler } from '@fixture/pilot';
 * const handler: ModelHandler<number> = value => value;
 * ```
 */
export interface ModelHandler<TOuter> {
  /** Returns the enclosing model value.
   * @typeParam TOuter - Model declared by the enclosing handler.
   * @param value - Required model value.
   * @returns Original model value.
   */
  (value: TOuter): TOuter;
}
/** A callable alias retaining separate enclosing and call generics.
 * @docPage shared
 * @typeParam TOuter - Required enclosing record model.
 * @example
 * ```ts
 * import { ModelSelector } from '@fixture/pilot';
 * const select: ModelSelector<{ name: string }> = value => value;
 * ```
 */
export type ModelSelector<TOuter> =
  /** Returns a selected key from the enclosing record.
   * @typeParam TOuter - Record declared by the enclosing selector.
   * @typeParam TKey - Required key retained by this call.
   * @param value - Required selected key.
   * @returns Original selected key.
   */
  <TKey extends keyof TOuter>(value: TKey) => TKey;
