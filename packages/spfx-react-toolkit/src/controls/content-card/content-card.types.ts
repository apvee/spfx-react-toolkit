/**
 * @packageDocumentation
 * @docPage controls/content-card
 */
import type { CardProps } from '@fluentui/react-components';
import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';

/**
 * Native Fluent Card appearance. Defaults to filled.
 * @example
 * ```ts
 * import { ContentCardAppearance } from '@apvee/spfx-react-toolkit';
 * const appearance: ContentCardAppearance = 'outline';
 * ```
 */
export type ContentCardAppearance = NonNullable<CardProps['appearance']>;
/**
 * Owned outer heading level, independent of visual typography. Defaults to 3.
 * @example
 * ```ts
 * import { ContentCardHeadingLevel } from '@apvee/spfx-react-toolkit';
 * const headingLevel: ContentCardHeadingLevel = 2;
 * ```
 */
export type ContentCardHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/**
 * Titled content using native Card interaction and fixed-tag Fluent content slots.
 * Strings are literal. Objects introduce inner containers; their refs never replace owned shells.
 * Callbacks run during rendering and must forward composed props. JSX remains caller-owned.
 * Native focus, selection, size, checkbox, floatingAction, events and root ref are preserved.
 * Uses the host Fluent theme/Griffel renderer without an SPFx provider. Optional regions
 * are omitted for absent/empty shorthand; explicit object/callback slots request their containers.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { ContentCard } from '@apvee/spfx-react-toolkit';
 * <ContentCard title="Workspace" actions={<button>Open</button>}>Three members</ContentCard>
 * ```
 */
export interface ContentCardProps extends Omit<CardProps, 'title' | 'children' | 'as' | 'dangerouslySetInnerHTML' | 'orientation' | 'appearance'> {
  /**
   * Required meaningful phrasing content inside the owned heading; fixed inner span.
   * No nested headings. Meaningless static titles throw `TypeError`; callback/custom JSX
   * content remains the caller's accessibility responsibility. Literal strings clamp to two lines.
   */
  title: Exclude<NonNullable<ContentSlot<'span'>>, boolean> & (object | string | number);
  /**
   * Owned heading tag h1–h6; when omitted, defaults to h3.
   * @defaultValue 3
   */
  headingLevel?: ContentCardHeadingLevel;
  /**
   * Optional supporting div slot. Literal strings, including object children, have a
   * two-line limit; JSX/callbacks own presentation. When omitted, leaves no description region.
   */
  description?: ContentSlot<'div'>;
  /**
   * Decorative header span slot, hidden from assistive technology.
   * When omitted, this icon region is absent.
   */
  icon?: ContentSlot<'span'>;
  /**
   * Header div slot for explicit caller controls, independent of footer actions.
   * When omitted, this headerActions region is absent.
   */
  headerActions?: ContentSlot<'div'>;
  /**
   * Caller-owned body; strings remain literal and controls retain identity.
   * When omitted, no child content is added.
   */
  children?: React.ReactNode;
  /**
   * Supplemental footer div slot; does not imply a timestamp or navigation.
   * When omitted, this footer region is absent.
   */
  footer?: ContentSlot<'div'>;
  /**
   * Independent footer div slot for explicit caller controls; buttons own their props.
   * When omitted, this actions region is absent.
   */
  actions?: ContentSlot<'div'>;
  /**
   * Native Fluent appearance; when omitted, defaults to filled. Size defaults to medium.
   * @defaultValue 'filled'
   */
  appearance?: ContentCardAppearance;
  /**
   * Raw HTML is unsupported; provide React children or safe content slots.
   * Must be omitted or undefined in this branch.
   */
  dangerouslySetInnerHTML?: never;
}
