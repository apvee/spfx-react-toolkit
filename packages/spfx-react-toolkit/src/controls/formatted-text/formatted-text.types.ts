import type * as React from 'react';
import type * as typography from '../../helpers/styles/typography';

/** Existing toolkit typography recipes; each controls appearance independently of HTML semantics.
 * @example const variant: FormattedTextVariant = 'title2';
 */
export type FormattedTextVariant = keyof typeof typography;

/** Supported semantic roots. The default is span, regardless of typography.
 * @example const as: FormattedTextSemanticElement = 'h2';
 */
export type FormattedTextSemanticElement = 'span' | 'p' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6';

/** Native text attributes plus bounded inline Markdown and toolkit typography.
 * A Fluent theme supplies CSS variables; no SPFx provider is required.
 * @example
 * <FormattedText as="h2" variant="title2" maxLines={2}>Read the **updates**.</FormattedText>
 */
export interface FormattedTextProps extends Omit<React.HTMLAttributes<HTMLElement>, 'children' | 'color' | 'dangerouslySetInnerHTML'> {
  /** HTML root, default span. Does not depend on variant; natural display follows this element. */
  as?: FormattedTextSemanticElement;
  /** Existing typography recipe, default body1. Does not select an HTML element. */
  variant?: FormattedTextVariant;
  /** Fluent font family, default base. Numeric selects a family, not number formatting. */
  font?: 'base' | 'monospace' | 'numeric';
  /** Text alignment, default start; start/end follow writing direction. */
  align?: 'start' | 'center' | 'end' | 'justify';
  /** Parse each owned string as safe inline Markdown, default true. JSX stays opaque. */
  markdown?: boolean;
  /** Apply the toolkit foreground.subtle descriptor, default false. */
  subtle?: boolean;
  /** Wrap resolved content in a semantic mark, default false. */
  highlight?: boolean;
  /** Italic presentation, default false. */
  italic?: boolean;
  /** Underline presentation, default false; composes with strikethrough. */
  underline?: boolean;
  /** Strike-through presentation, default false; composes with underline. */
  strikethrough?: boolean;
  /** Permit normal wrapping, default true; overridden by truncate or valid maxLines. */
  wrap?: boolean;
  /** Request single-line ellipsis, default false; requires constrained width and suitable layout. */
  truncate?: boolean;
  /** Positive finite integer line limit. Applies block clamp layout even to span without changing its tag.
   * Uses a generated class, not inline style. Invalid values are ignored; takes precedence over truncate/wrap. Consumer style may override it. */
  maxLines?: number;
  /** Owned strings, arrays and fragments receive Markdown; arbitrary JSX retains identity and contents. */
  children?: React.ReactNode;
  /** Not supported: children are the sole content pathway, and raw Markdown HTML is literal text. */
  dangerouslySetInnerHTML?: never;
}
