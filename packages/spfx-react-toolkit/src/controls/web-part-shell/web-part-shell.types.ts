import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { Theme } from '@fluentui/react-theme';

/** Persisted title properties owned by WebPartShell. Extend with application properties.
 * Do not narrow these keys: the shell publishes ordinary strings and visibility values.
 * @example
 * ```ts
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * ```
 */
export interface IWebPartShellProperties {
  /** Plain searchable title, kept in sync with rich-title edits. Literal H2 fallback when
   * `titleHTML` is absent; metadata searchability requires `getWebPartShellPropertiesMetadata`. */
  title: string;
  /** Authoritative toolkit rich-title HTML; defined empty means an empty title.
   * Saved HTML is normalized to one paragraph or H1-H4 block with supported inline formatting.
   * Input and canonical output are limited to 65,536 UTF-8 bytes; parser nesting is limited
   * to depth 32. Invalid or oversized input falls back to the plain title without a
   * mount-time property rewrite. See the WebPartShell guide's Title authority and HTML
   * protocol section for accepted tags and semantic class markers.
   * Title edits publish normalized safe markup together with the plain title.
   * Colors are stable semantic IDs resolved from the current Fluent theme;
   * UI-language changes never translate author content or rewrite these IDs. */
  titleHTML?: string;
  /** Display the title; undefined behaves as true without a property write. */
  showTitle?: boolean;
}

/** Host surface policy: `auto` detects Teams or a single-web-part app page as `application`;
 * other hosts use `section`. `section` adds no background or automatic inset; `application`
 * uses the themed application surface. Omission defaults to `auto`.
 * @example
 * ```ts
 * const surface: WebPartShellSurface = 'section';
 * ```
 */
export type WebPartShellSurface = 'auto' | 'section' | 'application';
/** Logical Fluent inset: `none`, `small`, `medium` or `large`. `auto` (default) uses none
 * for section and responsive medium/large application padding around the 640px container breakpoint.
 * @example
 * ```ts
 * const padding: WebPartShellPadding = 'small';
 * ```
 */
export type WebPartShellPadding = 'auto' | 'none' | 'small' | 'medium' | 'large';

/** Actual SPFx web-part shell options and native div attributes. Composes the toolkit
 * SPFx and Fluent providers internally, waiting for ServiceScope readiness. The host
 * must supply its document's Griffel renderer. No editor engine types are public.
 * Title editor labels, color tooltips and recovery messages follow the instance's
 * `pageContext.cultureInfo.currentUICultureName`, with English fallback for unsupported
 * cultures. Theme-color identifiers and author content remain language-independent.
 * The editor offers 12 theme tones and 10 standard hues from foreground tokens
 * and fixed theme-derived mixes, without custom HEX colors. Coincident values
 * are deduplicated while preserving the selected semantic color.
 * Inline formatting applies to selected text, or to the entire nonempty title
 * when only a caret is present. Toolbar commands and Ctrl/Cmd+B/I/U preserve
 * the caret/range and provide individual undo boundaries. Empty-title inline formatting
 * affects subsequent typing without publishing an empty document.
 * @template TProperties Web part properties extending the shell's three owned keys.
 * @example
 * ```tsx
 * <WebPartShell instance={this} padding="small"><MyContent /></WebPartShell>
 * ```
 */
export interface WebPartShellProps<TProperties extends IWebPartShellProperties>
  extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /** Actual SPFx web part; supplies properties, services, lifecycle and owner document. */
  instance: BaseClientSideWebPart<TProperties>;
  /** Override host surface detection; default auto. */
  surface?: WebPartShellSurface;
  /** Override automatic surface spacing; default auto. */
  padding?: WebPartShellPadding;
  /** Explicit Fluent theme; takes precedence over ancestor and live host themes. */
  theme?: Theme;
  /** Direction override; otherwise ancestor Fluent direction or SPFx locale. */
  dir?: 'ltr' | 'rtl';
  /** Header actions as shorthand or div props/render slot under both providers. Shorthand retains its DOM; slot ref targets the inner div. Empty objects request a header; callbacks run only during rendering. Persisted title/editor and body identity are independent. */
  action?: ContentSlot<'div'>;
  /** Content whose mounted identity is preserved through title Read/Edit and same-instance
   * rerenders. A different `instanceId` resets the subtree; omission leaves an empty body. */
  children?: React.ReactNode;
  /** HTML injection is unavailable; the shell owns the title content pathway. */
  dangerouslySetInnerHTML?: never;
}
