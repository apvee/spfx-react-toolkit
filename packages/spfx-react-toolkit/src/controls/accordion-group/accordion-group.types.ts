/**
 * @packageDocumentation
 * @docPage controls/accordion-group
 */
import type { ContentSlot } from '../content-slot.types.internal';
import type { AccordionProps as FluentAccordionProps } from '@fluentui/react-components';
import type * as React from 'react';

/**
 * Semantic heading level, independent of typography.
 * @example
 * ```ts
 * import { AccordionGroupHeadingLevel } from '@apvee/spfx-react-toolkit';
 * const level: AccordionGroupHeadingLevel = 2;
 * ```
 */
export type AccordionGroupHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;
/**
 * Stable item identifier; use a unique string or number in each group.
 * @example
 * ```ts
 * import { AccordionGroupValue } from '@apvee/spfx-react-toolkit';
 * const value: AccordionGroupValue = 'profile';
 * ```
 */
export type AccordionGroupValue = string | number;

/**
 * Native div group attributes, optional header and Fluent-owned expansion. Uses the
 * host Fluent theme/renderer with no SPFx provider; application controls own form state.
 * Strings stay literal. Text wrappers apply toolkit presets; JSX children retain their props.
 * Named props accept shorthand, fixed-tag native props or children render callbacks.
 * Objects request inner containers; refs identify those containers, not owned headings or triggers.
 * Callbacks receive the fixed tag/composed props during rendering and must forward the props.
 * Empty static descriptions do not add accessible associations; raw HTML is unsupported.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { AccordionGroup } from '@apvee/spfx-react-toolkit';
 * <AccordionGroup title="Preferences"><AccordionGroup.Item value="profile" title="Profile">Content</AccordionGroup.Item></AccordionGroup>
 * ```
 */
export interface AccordionGroupProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'onToggle' | 'dangerouslySetInnerHTML'> {
  /**
   * Optional inner span slot (shorthand, native props or render callback); heading phrasing content; wrapper uses subtitle2 and primary foreground. No nested headings.
   * When omitted, this title region is absent.
   */
  title?: ContentSlot<'span'>;
  /**
   * Optional inner span slot for supporting phrasing content; wrapper uses caption1 and subtle foreground.
   * When omitted, this description region is absent.
   */
  description?: ContentSlot<'span'>;
  /**
   * Optional inner span slot for a decorative header icon in a fixed 24px slot; font/SVG children scale without cloning. Hidden from assistive technology.
   * When omitted, this icon region is absent.
   */
  icon?: ContentSlot<'span'>;
  /**
   * Optional inner div slot for caller-owned header actions.
   * When omitted, this action region is absent.
   */
  action?: ContentSlot<'div'>;
  /**
   * Root title wrapper heading level for strings and JSX; when omitted, defaults to 2.
   * @defaultValue 2
   */
  headingLevel?: AccordionGroupHeadingLevel;
  /**
   * Native body Card appearance; when omitted, defaults to `filled`.
   * `alternate` uses Fluent's `filled-alternative` surface; `outline` uses its outlined surface.
   * Appearance changes preserve mounted controls and their state.
   * @defaultValue 'filled'
   */
  appearance?: 'filled' | 'alternate' | 'outline';
  /**
   * Controlled open value or values. Update from `onToggle` to accept a requested change;
   * omit this prop for internal Fluent state initialized by `defaultOpenItems`.
   */
  openItems?: FluentAccordionProps<AccordionGroupValue>['openItems'];
  /**
   * Initial uncontrolled open value or values; when omitted, defaults to no open items. Ignored
   * when `openItems` is supplied; later changes do not reset uncontrolled state.
   * @defaultValue No open items in uncontrolled mode.
   */
  defaultOpenItems?: FluentAccordionProps<AccordionGroupValue>['defaultOpenItems'];
  /**
   * Called when Fluent requests expansion/collapse. Receives its mouse/keyboard event
   * and `{ value, openItems }`: the toggled identifier and complete requested open list.
   * Controlled consumers update `openItems`; returned promises are not awaited.
   * When omitted, Fluent still manages uncontrolled expansion, but no consumer toggle notification is delivered.
   */
  onToggle?: FluentAccordionProps<AccordionGroupValue>['onToggle'];
  /**
   * Permit several open panels. when omitted, defaults to false.
   * @defaultValue false
   */
  multiple?: boolean;
  /**
   * Permit the last open panel to close. when omitted, defaults to true.
   * @defaultValue true
   */
  collapsible?: boolean;
  /**
   * Caller-owned compound items or custom content, rendered in declared order.
   * When omitted, no child content is added.
   */
  children?: React.ReactNode;
  /**
   * Raw HTML is unsupported; children and content props are the content pathways.
   * Must be omitted or undefined in this branch.
   */
  dangerouslySetInnerHTML?: never;
}

/**
 * One composable disclosure; caller controls retain identity, labels and state.
 * Named span slots accept shorthand, native props and render callbacks; refs point inside the trigger.
 * Requires an `AccordionGroup` ancestor. The heading, button and panel associations
 * remain owned by the disclosure; item native attributes/ref target its outer div.
 * Do not place interactive links, controls or block markup in trigger content.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { AccordionGroup } from '@apvee/spfx-react-toolkit';
 * <AccordionGroup><AccordionGroup.Item value="profile" title="Profile">Content</AccordionGroup.Item></AccordionGroup>
 * ```
 */
export interface AccordionGroupItemProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /**
   * Optional inner span slot for literal phrasing content; uses body1Strong/primary.
   * Supply a meaningful title or description to name the trigger; omit interactive descendants.
   */
  title?: ContentSlot<'span'>;
  /**
   * Optional inner span slot for description phrasing content; span wrapper uses caption1 and subtle foreground.
   * When omitted, this description region is absent.
   */
  description?: ContentSlot<'span'>;
  /**
   * Optional inner span slot for a decorative icon in a fixed 20px slot; font/SVG children scale without cloning. Hidden from assistive technology.
   * When omitted, this icon region is absent.
   */
  icon?: ContentSlot<'span'>;
  /**
   * Unique stable value within the parent AccordionGroup.
   */
  value: AccordionGroupValue;
  /**
   * Trigger heading level, default 3; independent of text appearance.
   * When omitted, defaults to 3.
   * @defaultValue 3
   */
  headingLevel?: AccordionGroupHeadingLevel;
  /**
   * Disable the native trigger. when omitted, defaults to false.
   * @defaultValue false
   */
  disabled?: boolean;
  /**
   * Reset descendants after closing motion. when omitted, defaults to false (retained and inert).
   * Initially closed resettable content does not mount; a policy change while closed
   * applies on the next opening/closing cycle. Closing does not suspend retained effects.
   * @defaultValue false
   */
  unmountOnClose?: boolean;
  /**
   * Caller-owned panel content, retained by default.
   * When omitted, no child content is added.
   */
  children?: React.ReactNode;
  /**
   * Raw HTML is unsupported; children own panel/control content.
   * Must be omitted or undefined in this branch.
   */
  dangerouslySetInnerHTML?: never;
}
