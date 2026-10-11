import type * as React from 'react';

/** Semantic heading level, independent of typography.
 * @example const level: SettingsGroupHeadingLevel = 2; */
export type SettingsGroupHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** Native group root attributes and optional content; no form state or implicit labeling.
 * Strings stay literal. Text wrappers apply toolkit presets; JSX children retain their props.
 * @example <SettingsGroup title="Preferences"><SettingsGroup.Item title="Profile">Content</SettingsGroup.Item></SettingsGroup> */
export interface SettingsGroupProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /** Optional heading phrasing content; wrapper uses subtitle2 and primary foreground. No nested headings. */
  title?: React.ReactNode;
  /** Optional supporting phrasing content; wrapper uses caption1 and subtle foreground. */
  description?: React.ReactNode;
  /** Optional decorative header icon in a fixed 24px slot; font/SVG children scale without cloning. Hidden from assistive technology. */
  icon?: React.ReactNode;
  /** Optional caller-owned header actions. */
  action?: React.ReactNode;
  /** Root title wrapper heading level for strings and JSX; defaults to 2. */
  headingLevel?: SettingsGroupHeadingLevel;
  /** Fluent Card appearance; defaults to filled. */
  appearance?: 'filled' | 'filled-alternative';
  /** Caller-owned compound items or custom content, rendered in declared order. */
  children?: React.ReactNode;
  /** Raw HTML is unsupported; children and content props are the content pathways. */
  dangerouslySetInnerHTML?: never;
}

/** One composable settings row; caller controls retain identity, labels and state.
 * Opt into field composition to label one built-in Form control using row text.
 * @example <SettingsGroup.Item title="Profile">Content</SettingsGroup.Item> */
export interface SettingsGroupItemProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /** Use title/description as one built-in Form logical control's label/hint.
   * Requires a meaningful title and exactly one primary built-in adapter. Empty static strings,
   * intrinsic markup and Fragments are rejected; numeric title content remains supported.
   * Custom components must render meaningful accessible title content; their output is caller-owned.
   * Omit the control's label/hint;
   * validation stays in Form. Actions, nested rows and custom useField bindings are excluded.
   * Defaults to false; composite controls name their group instead of every inner control.
   * @example <SettingsGroup.Item field title="Email" description="Receive alerts"><Order.Switch name="email" /></SettingsGroup.Item>
   */
  field?: boolean;
  /** Optional title phrasing content; span wrapper uses body1Strong and primary foreground. */
  title?: React.ReactNode;
  /** Optional description phrasing content; span wrapper uses caption1 and subtle foreground. */
  description?: React.ReactNode;
  /** Optional decorative icon in a fixed 20px slot; font/SVG children scale without cloning. Hidden from assistive technology. */
  icon?: React.ReactNode;
  /** Optional secondary caller-owned action, outside the primary controls. */
  action?: React.ReactNode;
  /** Primary controls; field mode supplies one built-in Form control with row naming. */
  children?: React.ReactNode;
  /** Raw HTML is unsupported; children own panel/control content. */
  dangerouslySetInnerHTML?: never;
}
