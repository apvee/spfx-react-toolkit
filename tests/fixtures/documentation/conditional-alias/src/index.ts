/** Guards object attributes while preserving nonobject shorthand.
 * @typeParam Props - Original object or shorthand.
 * @typeParam Owned - Attributes reserved by the outer owner.
 * @example
 * ```ts
 * import type { Surface } from '@fixture/conditional-alias';
 * const value: Surface<{ text: string }, 'owned'> = { text: 'Summary' };
 * ```
 */
type GuardedObject<Props, Owned extends PropertyKey> = Props extends object
  ? Omit<Props, Owned> & {
      /** Raw HTML must be omitted or undefined. */
      rawHtml?: never;
      /** Callable objects must be omitted or undefined. */
      call?: never;
      /** Nested options must be omitted or undefined. */
      options?: {
        /** Required nested label. */
        label: string;
      };
      /** Callback must be omitted or undefined.
       * @typeParam Item - Callback-local input.
       * @param value - Required callback input.
       * @returns Original callback input.
       */
      callback?: <Item extends Props>(value: Item) => Item;
    } & { [Key in Owned]?: never }
  : Props;

/** Allows guarded objects and numeric shorthand.
 * @docPage surface
 * @typeParam Props - Native object props.
 * @typeParam Owned - Reserved attribute keys.
 * @example
 * ```ts
 * import type { Surface } from '@fixture/conditional-alias';
 * const value: Surface<{ text: string }, 'owned'> = 1;
 * ```
 */
export type Surface<Props, Owned extends PropertyKey> =
  | GuardedObject<Props, Owned>
  | (GuardedObject<Props, Owned> & { [Key in `data-${string}`]?: string })
  | number;

/** Resolved object branch.
 * @docPage surface
 * @example
 * ```ts
 * import type { ObjectSurface } from '@fixture/conditional-alias';
 * const value: ObjectSurface = { text: 'Summary' };
 * ```
 */
export type ObjectSurface = GuardedObject<{
  /** Required original text. */
  text: string;
  /** Reserved attribute before guarding. */
  owned: string;
}, 'owned'>;

/** Resolved nonobject branch.
 * @docPage surface
 * @example
 * ```ts
 * import type { ShorthandSurface } from '@fixture/conditional-alias';
 * const value: ShorthandSurface = 'Summary';
 * ```
 */
export type ShorthandSurface = GuardedObject<string, 'owned'>;
