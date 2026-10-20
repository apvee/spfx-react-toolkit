import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';

/** Pair orientation; auto stacks without a named container and aligns at 480px.
 * @example
 * ```ts
 * const layout: DescriptionListLayout = 'horizontal';
 * ```
 */
export type DescriptionListLayout = 'auto' | 'horizontal' | 'vertical';

/** Immutable term/value pair for a single record's read-only metadata. The list preserves
 * input order and uses `id` as a React key; values remain caller-formatted literal content.
 * @example
 * ```ts
 * const owner: DescriptionListEntry = { id: 'owner', label: 'Owner', value: 'Ada' };
 * ```
 */
export interface DescriptionListEntry {
  /** Stable unique identity within the list; React key, not a generated DOM ID. */
  readonly id: string | number;
  /** Meaningful term: scalar, JSX or Fluent div-slot props/render callback. Boolean conditions
   * are absent; strings stay literal. Only as="div" is supported; raw HTML is unsupported. */
  readonly label: NonNullable<ContentSlot<'div'>>;
  /** Associated value inside a native dd. Undefined/boolean leaves it empty; Fluent legacy null
   * remains supported. Explicit props request a div even with empty children. Slot refs target this div. */
  readonly value: ContentSlot<'div'>;
}

/** Shared native dl attributes and pair layout; items and compound children are supplied
 * by the mutually exclusive public DescriptionListProps branches. Raw HTML is unavailable.
 * @example
 * ```ts
 * const layout: DescriptionListBaseProps = { layout: 'vertical', 'aria-label': 'Record details' };
 * ```
 */
type DescriptionListBaseProps = Omit<React.ComponentPropsWithoutRef<'dl'>, 'children' | 'dangerouslySetInnerHTML'> & {
  /** Default auto: stacked below 480px of the nearest apvee-sx container, horizontal at/above it.
   * Without that ancestor or CSS container-query support, auto stays stacked. */
  layout?: DescriptionListLayout;
  /** Unsupported raw-HTML pathway; use items or compound children. */
  dangerouslySetInnerHTML?: never;
};

/** Native dl attributes and mutually exclusive data or compound content.
 * Uses the host's Fluent theme and Griffel renderer, with no SPFx provider requirement.
 * @example
 * ```tsx
 * <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada' }]} />
 * ```
 */
export type DescriptionListProps = DescriptionListBaseProps & (
  | {
      /** Ordered immutable pairs with unique IDs. Empty arrays render nothing. */
      items: readonly DescriptionListEntry[];
      /** Unavailable when items supplies the content. */
      children?: never;
    }
  | {
      /** Unavailable when compound children supply the content. */
      items?: never;
      /** Compound Items, fragments and conditions in declaration order.
       * Absent/boolean children render nothing; an empty fragment may leave an empty dl. */
      children?: React.ReactNode;
    }
);

/** One compound pair, requiring DescriptionList and inheriting its layout.
 * Native attributes/ref target the grouping div, which has no box in horizontal layout.
 * Content slots remain inside invariant native dt/dd cells; caller JSX keeps its identity.
 * @example
 * ```tsx
 * <DescriptionList><DescriptionList.Item label="Owner">Ada</DescriptionList.Item></DescriptionList>
 * ```
 */
export interface DescriptionListItemProps extends Omit<React.ComponentPropsWithoutRef<'div'>, 'children' | 'dangerouslySetInnerHTML'> {
  /** Meaningful term as scalar, JSX or Fluent div-slot props/render callback; literal strings are not parsed.
   * as may only repeat "div". Consumer attributes/ref target the inner div, never the native dt. */
  label: NonNullable<ContentSlot<'div'>>;
  /** Associated value slot. Undefined/boolean/legacy Fluent null leaves an empty dd. Strings/JSX
   * retain their div wrapper; empty props objects request a div. Callbacks receive the fixed div
   * and composed props during rendering, without prior execution to determine presence. */
  children: ContentSlot<'div'>;
  /** Unsupported raw-HTML pathway; use label and children slots. */
  dangerouslySetInnerHTML?: never;
}
