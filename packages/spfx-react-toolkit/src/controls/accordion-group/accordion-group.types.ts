import type { AccordionProps as FluentAccordionProps } from '@fluentui/react-components';
import type * as React from 'react';

/** Semantic heading level, independent of typography.
 * @example const level: AccordionGroupHeadingLevel = 2; */
export type AccordionGroupHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;
/** Stable item identifier; use a unique string or number in each group.
 * @example const value: AccordionGroupValue = 'profile'; */
export type AccordionGroupValue = string | number;

/** Native group root attributes and optional content; no form state or implicit labeling.
 * Strings stay literal. Text wrappers apply toolkit presets; JSX children retain their props.
 * @example <AccordionGroup title="Preferences"><AccordionGroup.Item value="profile" title="Profile">Content</AccordionGroup.Item></AccordionGroup> */
export interface AccordionGroupProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'onToggle' | 'dangerouslySetInnerHTML'> {
  /** Optional heading phrasing content; wrapper uses subtitle2 and primary foreground. No nested headings. */
  title?: React.ReactNode;
  /** Optional supporting phrasing content; wrapper uses caption1 and subtle foreground. */
  description?: React.ReactNode;
  /** Optional decorative header icon in a fixed 24px slot; font/SVG children scale without cloning. Hidden from assistive technology. */
  icon?: React.ReactNode;
  /** Optional caller-owned header actions. */
  action?: React.ReactNode;
  /** Root title wrapper heading level for strings and JSX; defaults to 2. */
  headingLevel?: AccordionGroupHeadingLevel;
  /** Fluent Card appearance; defaults to filled. */
  appearance?: 'filled' | 'filled-alternative';
  /** Controlled expansion; the application remains authoritative. */
  openItems?: FluentAccordionProps<AccordionGroupValue>['openItems'];
  /** Initial uncontrolled expansion; defaults to no open items. */
  defaultOpenItems?: FluentAccordionProps<AccordionGroupValue>['defaultOpenItems'];
  /** Fluent event and requested value/openItems, unchanged. */
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
 * Do not place interactive links, controls or block markup in trigger content.
 * @example <AccordionGroup.Item value="profile" title="Profile">Content</AccordionGroup.Item> */
export interface AccordionGroupItemProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /** Optional title phrasing content; span wrapper uses body1Strong and primary foreground. */
  title?: React.ReactNode;
  /** Optional description phrasing content; span wrapper uses caption1 and subtle foreground. */
  description?: React.ReactNode;
  /** Optional decorative icon in a fixed 20px slot; font/SVG children scale without cloning. Hidden from assistive technology. */
  icon?: React.ReactNode;
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
