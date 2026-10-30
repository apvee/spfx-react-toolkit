import type { ContentSlot } from '../content-slot.types.internal';
import type { AccordionProps as FluentAccordionProps } from '@fluentui/react-components';
import type * as React from 'react';

/** Semantic heading level, independent of typography.
 * @example
 * ```ts
 * const level: AccordionGroupHeadingLevel = 2;
 * ```
 */
export type AccordionGroupHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;
/** Stable item identifier; use a unique string or number in each group.
 * @example
 * ```ts
 * const value: AccordionGroupValue = 'profile';
 * ```
 */
export type AccordionGroupValue = string | number;

/** Native div group attributes, optional header and Fluent-owned expansion. Uses the
 * host Fluent theme/renderer with no SPFx provider; application controls own form state.
 * Strings stay literal. Text wrappers apply toolkit presets; JSX children retain their props.
 * Named props accept shorthand, fixed-tag native props or children render callbacks.
 * Objects request inner containers; refs identify those containers, not owned headings or triggers.
 * Callbacks receive the fixed tag/composed props during rendering and must forward the props.
 * Empty static descriptions do not add accessible associations; raw HTML is unsupported.
 * @example
 * ```tsx
 * <AccordionGroup title="Preferences"><AccordionGroup.Item value="profile" title="Profile">Content</AccordionGroup.Item></AccordionGroup>
 * ```
 */
export interface AccordionGroupProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'onToggle' | 'dangerouslySetInnerHTML'> {
  /** Optional inner span slot (shorthand, native props or render callback); heading phrasing content; wrapper uses subtitle2 and primary foreground. No nested headings. */
  title?: ContentSlot<'span'>;
  /** Optional inner span slot for supporting phrasing content; wrapper uses caption1 and subtle foreground. */
  description?: ContentSlot<'span'>;
  /** Optional inner span slot for a decorative header icon in a fixed 24px slot; font/SVG children scale without cloning. Hidden from assistive technology. */
  icon?: ContentSlot<'span'>;
  /** Optional inner div slot for caller-owned header actions. */
  action?: ContentSlot<'div'>;
  /** Root title wrapper heading level for strings and JSX; defaults to 2. */
  headingLevel?: AccordionGroupHeadingLevel;
  /** Native body Card appearance; defaults to `filled`.
   * `alternate` uses Fluent's `filled-alternative` surface; `outline` uses its outlined surface.
   * Appearance changes preserve mounted controls and their state. */
  appearance?: 'filled' | 'alternate' | 'outline';
  /** Controlled open value or values. Update from `onToggle` to accept a requested change;
   * omit this prop for internal Fluent state initialized by `defaultOpenItems`. */
  openItems?: FluentAccordionProps<AccordionGroupValue>['openItems'];
  /** Initial uncontrolled open value or values; defaults to no open items. Ignored
   * when `openItems` is supplied; later changes do not reset uncontrolled state. */
  defaultOpenItems?: FluentAccordionProps<AccordionGroupValue>['defaultOpenItems'];
  /** Called when Fluent requests expansion/collapse. Receives its mouse/keyboard event
   * and `{ value, openItems }`: the toggled identifier and complete requested open list.
   * Controlled consumers update `openItems`; returned promises are not awaited. */
  onToggle?: FluentAccordionProps<AccordionGroupValue>['onToggle'];
  /** Permit several open panels. Defaults to false. */
  multiple?: boolean;
  /** Permit the last open panel to close. Defaults to true. */
  collapsible?: boolean;
  /** Caller-owned compound items or custom content, rendered in declared order. */
  children?: React.ReactNode;
  /** Raw HTML is unsupported; children and content props are the content pathways. */
  dangerouslySetInnerHTML?: never;
}

/** One composable disclosure; caller controls retain identity, labels and state.
 * Named span slots accept shorthand, native props and render callbacks; refs point inside the trigger.
 * Requires an `AccordionGroup` ancestor. The heading, button and panel associations
 * remain owned by the disclosure; item native attributes/ref target its outer div.
 * Do not place interactive links, controls or block markup in trigger content.
 * @example
 * ```tsx
 * <AccordionGroup><AccordionGroup.Item value="profile" title="Profile">Content</AccordionGroup.Item></AccordionGroup>
 * ```
 */
export interface AccordionGroupItemProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /** Optional inner span slot for literal phrasing content; uses body1Strong/primary.
   * Supply a meaningful title or description to name the trigger; omit interactive descendants. */
  title?: ContentSlot<'span'>;
  /** Optional inner span slot for description phrasing content; span wrapper uses caption1 and subtle foreground. */
  description?: ContentSlot<'span'>;
  /** Optional inner span slot for a decorative icon in a fixed 20px slot; font/SVG children scale without cloning. Hidden from assistive technology. */
  icon?: ContentSlot<'span'>;
  /** Unique stable value within the parent AccordionGroup. */
  value: AccordionGroupValue;
  /** Trigger heading level, default 3; independent of text appearance. */
  headingLevel?: AccordionGroupHeadingLevel;
  /** Disable the native trigger. Defaults to false. */
  disabled?: boolean;
  /** Reset descendants after closing motion. Defaults to false (retained and inert).
   * Initially closed resettable content does not mount; a policy change while closed
   * applies on the next opening/closing cycle. Closing does not suspend retained effects. */
  unmountOnClose?: boolean;
  /** Caller-owned panel content, retained by default. */
  children?: React.ReactNode;
  /** Raw HTML is unsupported; children own panel/control content. */
  dangerouslySetInnerHTML?: never;
}
