/**
 * @packageDocumentation
 * @docPage controls/web-part-shell
 */
import type { EditableTextStringsOverride } from '../editable-text/editable-text.types';
import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { CardProps } from '@fluentui/react-components';
import type { Theme } from '@fluentui/react-theme';

/**
 * Persisted title properties owned by WebPartShell. Extend with application properties.
 * Do not narrow these keys: the shell publishes ordinary strings and visibility values.
 * @example
 * ```ts
 * import { IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * ```
 */
export interface IWebPartShellProperties {
  /**
   * Plain searchable title, kept in sync with rich-title edits. Literal H2 fallback when
   * `titleHTML` is absent; metadata searchability requires `getWebPartShellPropertiesMetadata`.
   */
  title: string;
  /**
   * Authoritative toolkit rich-title HTML; defined empty means an empty title.
   * Saved HTML is normalized to one paragraph or H1-H4 block with supported inline formatting.
   * Input and canonical output are limited to 65,536 UTF-8 bytes; parser nesting is limited
   * to depth 32. Invalid or oversized input falls back to the plain title without a
   * mount-time property rewrite. See the WebPartShell guide's Title authority and HTML
   * protocol section for accepted tags and semantic class markers.
   * Title edits publish normalized safe markup together with the plain title.
   * Colors are stable semantic IDs resolved from the current Fluent theme;
   * UI-language changes never translate author content or rewrite these IDs.
   * When omitted, the plain title property supplies the initial/read title.
   */
  titleHTML?: string;
  /**
   * Display the complete title/action header; undefined behaves as true without a property write.
   * False unmounts the header and suppresses action render callbacks.
   * @defaultValue true
   */
  showTitle?: boolean;
}

/**
 * Host surface policy: `auto` detects Teams or a single-web-part app page as `application`;
 * other hosts use `section`. Both wrappers remain transparent. `section` adds no
 * automatic inset; `application` enables responsive automatic outer padding.
 * Card appearance controls the optional visual surface. When omitted, defaults to `auto`.
 * @example
 * ```ts
 * import { WebPartShellSurface } from '@apvee/spfx-react-toolkit';
 * const surface: WebPartShellSurface = 'section';
 * ```
 */
export type WebPartShellSurface = 'auto' | 'section' | 'application';
/**
 * Logical Fluent inset: `none`, `small`, `medium` or `large`. `auto` (default) uses none
 * for section and responsive medium/large application padding around the 640px container breakpoint.
 * @example
 * ```ts
 * import { WebPartShellPadding } from '@apvee/spfx-react-toolkit';
 * const padding: WebPartShellPadding = 'small';
 * ```
 */
export type WebPartShellPadding = 'auto' | 'none' | 'small' | 'medium' | 'large';

/**
 * Native Fluent Card appearance, defaulting to `subtle`. Subtle renders no Card;
 * filled, filled-alternative and outline use a vertical Card with focusMode off.
 * Card size defaults to medium and can be selected independently.
 * Card spacing and visual treatment follow the host Fluent theme.
 * @example
 * ```ts
 * import { WebPartShellAppearance } from '@apvee/spfx-react-toolkit';
 * const appearance: WebPartShellAppearance = 'outline';
 * ```
 */
export type WebPartShellAppearance = NonNullable<CardProps['appearance']>;

/**
 * Native Fluent Card size: small, medium or large. When omitted, defaults to medium.
 * Controls the Card's native padding, gap and radius; outer surface padding is independent.
 * Subtle appearance ignores size because it renders no Card. Size updates preserve
 * the Card, header/editor and body identity without writing persisted properties.
 * @example
 * ```ts
 * import { WebPartShellSize } from '@apvee/spfx-react-toolkit';
 * const size: WebPartShellSize = 'small';
 * ```
 */
export type WebPartShellSize = NonNullable<CardProps['size']>;

/**
 * Placement of the complete title/action header relative to a rendered Card.
 * Defaults to outside. Subtle ignores this option because it has no Card.
 * Changing placement can remount the header while preserving body state.
 * @example
 * ```ts
 * import { WebPartShellTitlePosition } from '@apvee/spfx-react-toolkit';
 * const titlePosition: WebPartShellTitlePosition = 'inside';
 * ```
 */
export type WebPartShellTitlePosition = 'outside' | 'inside';

/**
 * Shell-owned title naming, independent of editor command dictionaries. Native
 * SPFx resources supply omitted override fields; these labels never alter author
 * title content or persisted properties. Empty placeholders suppress the hint.
 * An empty editor label still requires a valid accessible name in Edit.
 * @example
 * ```ts
 * import { WebPartShellStrings } from '@apvee/spfx-react-toolkit';
 * const labels: WebPartShellStrings = { titleEditorLabel: 'Section title', titlePlaceholder: 'Add a heading' };
 * ```
 */
export interface WebPartShellStrings {
  /**
   * Accessible name passed to the actual title editor textbox.
   */
  readonly titleEditorLabel: string;
  /**
   * Literal empty title hint, never serialized or used as the accessible name.
   */
  readonly titlePlaceholder: string;
}

/**
 * Actual SPFx web-part shell options and native div attributes. Composes the toolkit
 * SPFx and Fluent providers internally, waiting for ServiceScope readiness. The host
 * must supply its document's Griffel renderer. No editor engine types are public.
 * Native SPFx resources select title/editor labels. Instance dictionaries can
 * override shell naming and editor commands independently. Theme-color identifiers
 * and author content remain language-independent.
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
 * import * as React from 'react';
 * import { IWebPartShellProperties, WebPartShell } from '@apvee/spfx-react-toolkit';
 * import { Button } from '@fluentui/react-components';
 * import * as ReactDom from 'react-dom';
 * import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * interface Properties extends IWebPartShellProperties { listId: string; }
 * export class ExampleWebPart extends BaseClientSideWebPart<Properties> {
 *   public render(): void {
 *     ReactDom.render(<WebPartShell instance={this} appearance="filled" size="small" titlePosition="inside"
 *       action={<Button onClick={() => this.render()}>Refresh</Button>}>
 *       <p>Workspace content</p>
 *     </WebPartShell>, this.domElement);
 *   }
 *   protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
 * }
 * ```
 */
export interface WebPartShellProps<TProperties extends IWebPartShellProperties>
  extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  /**
   * Actual SPFx web part; supplies properties, services, lifecycle and owner document.
   */
  instance: BaseClientSideWebPart<TProperties>;
  /**
   * Override host surface detection; default auto.
   * When omitted, defaults to auto.
   * @defaultValue 'auto'
   */
  surface?: WebPartShellSurface;
  /**
   * Override outer surface spacing; default auto. Native Card owns its inner spacing.
   * When omitted, defaults to auto.
   * @defaultValue 'auto'
   */
  padding?: WebPartShellPadding;
  /**
   * Optional native Card presentation; default subtle renders no Card. These are React
   * options, not persisted shell properties. Non-subtle changes preserve Card/body identity;
   * switching between subtle and a Card can remount body content. Card owns inner spacing.
   * When omitted, defaults to subtle renders no Card.
   * @defaultValue 'subtle'
   */
  appearance?: WebPartShellAppearance;
  /**
   * Native Card size; omitted/undefined defaults to medium. Small, medium and large
   * control native Card padding, gap and radius only, independently of outer padding.
   * Subtle ignores size. Updates preserve Card, header/editor and body state without
   * writing persisted properties.
   * @defaultValue 'medium'
   */
  size?: WebPartShellSize;
  /**
   * Complete header placement; default outside. Inside moves title and actions into the
   * Card; subtle ignores placement. Placement changes can remount the header only.
   * When omitted, defaults to outside.
   * @defaultValue 'outside'
   */
  titlePosition?: WebPartShellTitlePosition;
  /**
   * Explicit Fluent theme; takes precedence over ancestor and live host themes.
   * When omitted, inherits the nearest Fluent theme, then the live SPFx host theme.
   */
  theme?: Theme;
  /**
   * Direction override; otherwise ancestor Fluent direction or SPFx locale.
   * When omitted, inherits Fluent direction or the SPFx locale direction.
   */
  dir?: 'ltr' | 'rtl';
  /**
   * Shell title naming overrides. Omitted/undefined fields inherit native SPFx resources;
   * empty hints suppress the placeholder. Changes preserve title draft/history without writes.
   */
  strings?: Partial<WebPartShellStrings>;
  /**
   * Editor command, color, loading and recovery labels, merged with native resources.
   * Independent of shell naming: titlePlaceholder/titleEditorLabel remain owned by strings.
   * Replacing this dictionary preserves the live title editor and never writes properties.
   * When omitted, native editor resources supply all labels.
   */
  titleEditorStrings?: EditableTextStringsOverride;
  /**
   * Header actions as shorthand or div props/render slot under both providers.
   * Shorthand retains its DOM; slot ref targets the inner div. Visible-header objects
   * request the inner action container; callbacks run only when showTitle is true.
   * False showTitle unmounts title and actions together. Persisted title/editor and
   * body identity are independent.
   * When omitted, no header action content is rendered.
   */
  action?: ContentSlot<'div'>;
  /**
   * Content whose mounted identity is preserved through title Read/Edit and same-instance
   * rerenders, header visibility/placement, Card size and non-subtle appearance updates. Switching between
   * subtle and a Card can remount children. A different `instanceId` resets the subtree;
   * when omitted, leaves an empty body.
   */
  children?: React.ReactNode;
  /**
   * HTML injection is unavailable; the shell owns the title content pathway.
   * Must be omitted or undefined in this branch.
   */
  dangerouslySetInnerHTML?: never;
}
