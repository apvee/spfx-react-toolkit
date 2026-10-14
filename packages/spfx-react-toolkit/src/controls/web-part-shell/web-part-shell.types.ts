import type * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { Theme } from '@fluentui/react-theme';

/** Persisted title properties owned by WebPartShell. Extend with application properties.
 * Do not narrow these keys: the shell publishes ordinary strings and visibility values.
 * @example interface Properties extends IWebPartShellProperties { listId: string; }
 */
export interface IWebPartShellProperties {
  /** Searchable plain title. Literal H2 fallback when titleHTML is absent. */
  title: string;
  /** Authoritative toolkit rich title HTML; defined empty means an empty title.
   * Colors are stable semantic IDs resolved from the current Fluent theme;
   * UI-language changes never translate author content or rewrite these IDs. */
  titleHTML?: string;
  /** Display the title; undefined behaves as true without a property write. */
  showTitle?: boolean;
}

/** Automatic host detection, an inset-free section, or an application surface. */
export type WebPartShellSurface = 'auto' | 'section' | 'application';
/** Logical Fluent spacing; auto is none for section and responsive medium/large for application. */
export type WebPartShellPadding = 'auto' | 'none' | 'small' | 'medium' | 'large';

/** SPFx shell options and native div attributes. No editor engine types are public.
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
 * @example <WebPartShell instance={this} padding="small"><MyContent /></WebPartShell>
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
  /** Caller-owned header actions, rendered unchanged under both providers. */
  action?: React.ReactNode;
  /** Content whose mounted identity is preserved through title Read/Edit changes. */
  children?: React.ReactNode;
  /** HTML injection is unavailable; the shell owns the title content pathway. */
  dangerouslySetInnerHTML?: never;
}
