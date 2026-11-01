/**
 * @packageDocumentation
 * @docPage controls/settings-group
 */
import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';

/**
 * Semantic heading level, independent of typography.
 * @example
 * ```ts
 * import { SettingsGroupHeadingLevel } from '@apvee/spfx-react-toolkit';
 * const level: SettingsGroupHeadingLevel = 2;
 * ```
 */
export type SettingsGroupHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/**
 * Native div group attributes and optional header/body content; application controls own
 * values and validation. Uses the host Fluent theme/renderer with no SPFx provider.
 * Strings stay literal. Text wrappers apply toolkit presets; JSX children retain their props.
 * Named props accept shorthand, fixed-tag native props or children render callbacks.
 * Objects request inner containers; refs identify those containers, not owned headings/Labels.
 * Callbacks receive the fixed tag/composed props during rendering and must forward the props.
 * Empty static descriptions do not add accessible associations; raw HTML is unsupported.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SettingsGroup } from '@apvee/spfx-react-toolkit';
 * <SettingsGroup title="Preferences"><SettingsGroup.Item title="Profile">Content</SettingsGroup.Item></SettingsGroup>
 * ```
 */
export interface SettingsGroupProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /**
   * Optional inner span slot (shorthand, native props or render callback); heading phrasing content; wrapper uses subtitle2 and primary foreground. No nested headings; when omitted, removes the title region.
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
   * Optional inner div slot for caller-owned header actions; when omitted, leaves no action region.
   */
  action?: ContentSlot<'div'>;
  /**
   * Root title wrapper heading level for strings and JSX; when omitted, defaults to 2.
   * @defaultValue 2
   */
  headingLevel?: SettingsGroupHeadingLevel;
  /**
   * Native body Card appearance; when omitted, defaults to `filled`.
   * `alternate` uses Fluent's `filled-alternative` surface; `outline` uses its outlined surface.
   * Appearance changes preserve mounted controls and their state.
   * @defaultValue 'filled'
   */
  appearance?: 'filled' | 'alternate' | 'outline';
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
 * One composable settings row; caller controls retain identity, labels and state.
 * Named slots accept shorthand, fixed span props (div for action) and render callbacks.
 * Inner refs/classes/events are forwarded; semantic Label/IDs remain owned by the row.
 * Opt into field composition to label one built-in Form control using row text.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SettingsGroup } from '@apvee/spfx-react-toolkit';
 * <SettingsGroup><SettingsGroup.Item title="Profile">Content</SettingsGroup.Item></SettingsGroup>
 * ```
 */
export interface SettingsGroupItemProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /**
   * Use title/description as one built-in Form logical control's label/hint.
   * Requires a meaningful title and exactly one primary built-in adapter inside its Form root.
   * Empty/whitespace strings and empty intrinsic markup or Fragments are rejected as titles;
   * populated markup/Fragments and numeric title content remain supported.
   * Custom components and render callbacks must render meaningful accessible title content; their output is caller-owned.
   * Omit the control's label/hint;
   * validation stays in Form. Actions, nested rows and custom useField bindings are excluded.
   * Defaults to false; composite controls name their group instead of every inner control.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { SettingsGroup, createForm } from '@apvee/spfx-react-toolkit';
   * import { Switch } from '@fluentui/react-components';
   * const Order = createForm<{ email: boolean }>();
   * <Order.Root defaultValues={{ email: false }} onSubmit={values => { console.log(values.email); }}><SettingsGroup><SettingsGroup.Item field title="Email" description="Receive alerts"><Order.Switch name="email" /></SettingsGroup.Item></SettingsGroup></Order.Root>
   * ```
   *
   * @defaultValue false
   */
  field?: boolean;
  /**
   * Optional inner span slot (shorthand, native props or render callback); title phrasing content; span wrapper uses body1Strong and primary foreground.
   * When field is omitted or false, omitting title removes this region. Field mode
   * instead requires a meaningful title and throws when it is omitted or empty.
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
   * Optional inner div slot for a secondary caller-owned action, outside the primary controls.
   * When omitted, this action region is absent.
   */
  action?: ContentSlot<'div'>;
  /**
   * Primary controls; field mode supplies one built-in Form control with row naming.
   * When omitted, the row has no primary control content; field mode requires the consumer to supply its Form adapter.
   */
  children?: React.ReactNode;
  /**
   * Raw HTML is unsupported; children own panel/control content.
   * Must be omitted or undefined in this branch.
   */
  dangerouslySetInnerHTML?: never;
}
