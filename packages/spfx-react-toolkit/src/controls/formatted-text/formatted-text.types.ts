import type * as React from 'react';
import type * as typography from '../../helpers/styles/typography';

/** Existing toolkit typography recipes; each controls appearance independently of HTML semantics.
 * @example
 * ```ts
 * const variant: FormattedTextVariant = 'title2';
 * ```
 */
export type FormattedTextVariant = keyof typeof typography;

/** Supported semantic roots. The default is span, regardless of typography.
 * @example
 * ```ts
 * const as: FormattedTextSemanticElement = 'h2';
 * ```
 */
export type FormattedTextSemanticElement = 'span' | 'p' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6';

/** Native text attributes plus bounded inline Markdown and toolkit typography.
 * A Fluent theme supplies CSS variables; no SPFx provider is required.
 * Group mixed Markdown strings and JSX actions with `React.Fragment`; strings inside other
 * JSX elements, including `<span>`, remain literal.
 * @example
 * ```tsx
 * <FormattedText as="h2" variant="title2" maxLines={2}>Read the **updates**.</FormattedText>
 * ```
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
  /** Parse owned strings, including arrays and `React.Fragment`, as safe inline Markdown; defaults
   * to `true`. Other JSX elements keep their children unchanged: `<span>{'**text**'}</span>`
   * stays literal even when enabled. Use a fragment to group Markdown strings and JSX actions.
   * Supported emphasis, code and links do not enable raw HTML or block Markdown; unsafe links
   * retain their text. Each owned string above 64 KiB UTF-8 falls back literally before parsing.
   * More than 2,000 original AST nodes or depth above 32 (root depth 1) also returns that entire
   * string literally. Node/depth checks occur after parsing and do not bound the parser's work.
   * See the FormattedText guide's Markdown and mixed JSX section for the full grammar and limits. */
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
   * Omission adds no line clamp. Invalid values are ignored; takes precedence over truncate/wrap. Consumer style may override it. */
  maxLines?: number;
  /** Owned strings, arrays and `React.Fragment` receive Markdown when enabled. Use a fragment
   * to group mixed strings and JSX actions. Other JSX elements, including `<span>`, retain their
   * identity, props, refs and contents; their descendants are not parsed. Markdown limits and
   * complete literal fallback apply separately to each owned string, as described by markdown. */
  children?: React.ReactNode;
  /** Not supported: children are the sole content pathway, and raw Markdown HTML is literal text. */
  dangerouslySetInnerHTML?: never;
}
