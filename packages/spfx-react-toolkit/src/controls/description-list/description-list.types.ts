import type * as React from 'react';
import type { Slot } from '@fluentui/react-components';

/** Strip the raw-HTML pathway from Fluent object slots without changing shorthand content. */
type SafeSlot<Value> = Value extends { dangerouslySetInnerHTML?: unknown }
  ? Omit<Value, 'dangerouslySetInnerHTML'> & { dangerouslySetInnerHTML?: never }
  : Value;
type ContentSlot = SafeSlot<Slot<'div'>> | undefined;

/** Pair orientation; auto stacks without a named container and aligns at 480px.
 * @example const layout: DescriptionListLayout = 'horizontal'; */
export type DescriptionListLayout = 'auto' | 'horizontal' | 'vertical';

/** Immutable term/value pair for a single record's read-only metadata.
 * @example const owner: DescriptionListEntry = { id: 'owner', label: 'Owner', value: 'Ada' }; */
export interface DescriptionListEntry {
  /** Stable unique identity within the list; React key, not a generated DOM ID. */
  readonly id: string | number;
  /** Meaningful term: scalar, JSX or Fluent div-slot props/render callback. Raw HTML is unsupported. */
  readonly label: NonNullable<ContentSlot>;
  /** Associated value inside a native dd. Undefined leaves it empty; Fluent legacy null slots also remain supported. */
  readonly value: ContentSlot;
}

type DescriptionListBaseProps = Omit<React.ComponentPropsWithoutRef<'dl'>, 'children' | 'dangerouslySetInnerHTML'> & {
  /** Default auto: stacked below 480px of the nearest apvee-sx container, horizontal at/above it.
   * Without that ancestor or CSS container-query support, auto stays stacked. */
  layout?: DescriptionListLayout;
  /** Unsupported raw-HTML pathway; use items or compound children. */
  dangerouslySetInnerHTML?: never;
};

/** Native dl attributes and mutually exclusive data or compound content.
 * Uses the host's Fluent theme and Griffel renderer, with no SPFx provider requirement.
 * @example <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada' }]} /> */
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
 * @example <DescriptionList.Item label="Owner">Ada</DescriptionList.Item> */
export interface DescriptionListItemProps extends Omit<React.ComponentPropsWithoutRef<'div'>, 'children' | 'dangerouslySetInnerHTML'> {
  /** Meaningful term as scalar, JSX or Fluent div-slot props/render callback; literal strings are not parsed. */
  label: NonNullable<ContentSlot>;
  /** Associated value slot. Undefined is the default absence; legacy Fluent null leaves an empty dd. */
  children: ContentSlot;
  /** Unsupported raw-HTML pathway; use label and children slots. */
  dangerouslySetInnerHTML?: never;
}
