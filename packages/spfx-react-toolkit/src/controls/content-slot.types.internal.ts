import type { ExtractSlotProps, Slot } from '@fluentui/react-utilities';

/**
 * Declaration helper removing raw HTML and control-owned attributes from Fluent slot objects.
 * `Props` selects the native object branch; `Owned` names attributes whose semantics belong
 * to the outer control. Nonobject shorthand passes through unchanged.
 * @example
 * ```ts
 * import { ContentCardProps } from '@apvee/spfx-react-toolkit';
 * const content: ContentCardProps['title'] = { children: 'Summary' };
 * ```
 *
 * @typeParam Props - Native slot object props before control-owned attributes are removed.
 * @typeParam Owned - Outer-control attributes prohibited on inner slot objects.
 */
type SafeObjectSlot<Props, Owned extends PropertyKey> = Props extends object
  ? Omit<Props, 'dangerouslySetInnerHTML' | Owned> & {
      /**
       * Raw HTML is forbidden; safe React children are the only content pathway.
       * Must be omitted or undefined in this branch.
       */
      dangerouslySetInnerHTML?: never;
      /**
       * Functions belong in children; reject callable objects even with React 17 weak props.
       * Must be omitted or undefined in this branch.
       */
      call?: never;
    } & { [Key in Owned]?: never }
  : Props;

/**
 * Fluent slot object props with optional custom `data-*` attributes and ownership guards.
 * Retains the original object branch for React 17 JSX/ReactNode compatibility; `Props`
 * selects allowed attributes and `Owned` selects control-owned keys that cannot be supplied.
 * @example
 * ```ts
 * import { ContentCardProps } from '@apvee/spfx-react-toolkit';
 * const content: ContentCardProps['description'] = { children: 'Summary', 'data-region': 'summary' };
 * ```
 *
 * @typeParam Owned - Attributes owned by the outer control and forbidden on the inner slot.
 * @typeParam Props - Native slot object props before control-owned attributes are removed.
 */
type SafeSlotObjects<Props, Owned extends PropertyKey> =
  | SafeObjectSlot<Props, Owned>
  | (SafeObjectSlot<Props, Owned> & { [Key in `data-${string}`]?: string | number | boolean | undefined });

/**
 * Fixed-tag content accepted by named control regions: React shorthand, native slot props
 * or a props object with `children: (Component, props) => ReactNode`. `Tag` selects the
 * allowed native tag; `Owned` excludes outer-control attributes. Forward the composed
 * callback props to retain refs/classes/events; callbacks execute during rendering only.
 * Props objects request inner containers, while shorthand wrapping follows each control's
 * region policy. Booleans and undefined are absent; null remains Fluent legacy absence.
 * React 17 ReactNode variables can erase object restrictions; runtime guards remain necessary.
 * Associated declaration type only: import the control's public props rather than this private file.
 * @example
 * ```ts
 * import { ContentCardProps } from '@apvee/spfx-react-toolkit';
 * const title: ContentCardProps['title'] = { children: 'Summary', className: 'title' };
 * ```
 *
 * @typeParam Owned - Attributes owned by the outer control and forbidden on the inner slot.
 * @typeParam Tag - Fixed native element accepted by this control region.
 */
export type ContentSlot<Tag extends keyof JSX.IntrinsicElements, Owned extends PropertyKey = never> =
  | SafeSlotObjects<ExtractSlotProps<Slot<Tag>>, Owned>
  | Exclude<Slot<Tag>, ExtractSlotProps<Slot<Tag>>>
  | boolean
  | undefined;

/**
 * Named-region content with a primary native `Tag` and one documented `Alternative`.
 * `Owned` excludes outer-control attributes. Like ContentSlot, accepts shorthand or props
 * with static/rendered children; callback props/ref must be forwarded and absence uses
 * undefined/booleans (plus Fluent legacy null). The consumer chooses the alternative via `as`.
 * Only use the alternatives promised by the owning public control; no arbitrary tag replacement.
 * Associated declaration type only: consumers use the owning control's public props.
 * @example
 * ```ts
 * import { PreviewCardProps } from '@apvee/spfx-react-toolkit';
 * const logo: PreviewCardProps['logo'] = { as: 'img', src: '/logo.svg', alt: 'Brand' };
 * ```
 *
 * @typeParam Alternative - Explicitly permitted alternative native tag for this region.
 * @typeParam Owned - Attributes owned by the outer control and forbidden on the inner slot.
 * @typeParam Tag - Fixed native element accepted by this control region.
 */
export type AlternativeContentSlot<
  Tag extends keyof JSX.IntrinsicElements,
  Alternative extends keyof JSX.IntrinsicElements,
  Owned extends PropertyKey = never,
> =
  | SafeSlotObjects<ExtractSlotProps<Slot<Tag, Alternative>>, Owned>
  | Exclude<Slot<Tag, Alternative>, ExtractSlotProps<Slot<Tag, Alternative>>>
  | boolean
  | undefined;
