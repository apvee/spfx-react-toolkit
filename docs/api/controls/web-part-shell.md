# WebPartShell

`WebPartShell` composes the existing SPFx web part provider, a Fluent UI 9 provider, adaptive spacing and an inline editable title around an actual `BaseClientSideWebPart` instance. Children and header actions can use the toolkit hooks immediately after the instance's ServiceScope is ready. Extend `IWebPartShellProperties` with application properties; keep its three keys at their declared types.

## Usage

The following `.tsx` web part uses React 17, TypeScript 5.3 and SPFx 1.21 conventions. No external SPFx or Fluent provider is required for this shell.

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart, type IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { PropertyPaneToggle, type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { Button } from '@fluentui/react-components';
import { WebPartShell, getWebPartShellPropertiesMetadata,
  type IWebPartShellProperties, useSPFxProperties } from '@apvee/spfx-react-toolkit';

interface Properties extends IWebPartShellProperties {
  description: string;
}

function Content(): React.ReactElement {
  const { properties } = useSPFxProperties<Properties>();
  return <p>{properties?.description}</p>;
}

export default class ExampleWebPart extends BaseClientSideWebPart<Properties> {
  public render(): void {
    ReactDom.render(
      <WebPartShell instance={this} action={<Button onClick={() => this.render()}>Refresh</Button>}>
        <Content />
      </WebPartShell>,
      this.domElement
    );
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      ...getWebPartShellPropertiesMetadata(),
      description: { isSearchablePlainText: true }
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ groups: [{ groupName: 'Title', groupFields: [
      PropertyPaneToggle('showTitle', { label: 'Show title', checked: this.properties.showTitle ?? true })
    ] }] }] };
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

In the manifest's `preconfiguredEntries[].properties`, initialize application defaults and the plain title. Omit `titleHTML` to use the literal H2 fallback until an author edits it:

```json
{ "title": "Example", "showTitle": true, "description": "Application content" }
```

Authors edit the title on the canvas in SPFx Edit mode. The property pane toggle controls visibility; a second plain-title field would compete with authoritative rich HTML. Existing pages can leave `showTitle` undefined; the shell displays the title without writing a default.

## Public API

```typescript
interface IWebPartShellProperties {
  title: string;
  titleHTML?: string;
  showTitle?: boolean;
}
type WebPartShellSurface = 'auto' | 'section' | 'application';
type WebPartShellPadding = 'auto' | 'none' | 'small' | 'medium' | 'large';
interface WebPartShellProps<TProperties extends IWebPartShellProperties>
  extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'dangerouslySetInnerHTML'> {
  instance: BaseClientSideWebPart<TProperties>;
  surface?: WebPartShellSurface;
  padding?: WebPartShellPadding;
  theme?: Theme;
  dir?: 'ltr' | 'rtl';
  action?: React.ReactNode;
  children?: React.ReactNode;
  dangerouslySetInnerHTML?: never;
}
// Generic forwardRef component; ref points to the layout div below FluentProvider.
WebPartShell<TProperties extends IWebPartShellProperties>(
  props: WebPartShellProps<TProperties> & React.RefAttributes<HTMLDivElement>
): React.ReactElement;
function getWebPartShellPropertiesMetadata(): IWebPartPropertiesMetadata;
```

`Theme` comes from `@fluentui/react-theme`. `BaseClientSideWebPart` and `IWebPartPropertiesMetadata` come from `@microsoft/sp-webpart-base`.

| Member | Contract |
| --- | --- |
| `IWebPartShellProperties.title` | Searchable plain title; literal H2 fallback when rich HTML is absent or invalid. |
| `IWebPartShellProperties.titleHTML` | Authoritative toolkit HTML when defined, including the empty string. |
| `IWebPartShellProperties.showTitle` | Undefined means true. False hides title/editor without deleting saved values. |
| `WebPartShellProps.instance` | Actual SPFx instance, supplying properties, lifecycle, ServiceScope and owner document. |
| `WebPartShellProps.surface` | Defaults to auto; section/application override host detection. |
| `WebPartShellProps.padding` | Defaults to auto; explicit none/small/medium/large uses the existing Fluent spacing scale. |
| `WebPartShellProps.theme` | Explicit Fluent theme tokens override ancestor and host tokens. |
| `WebPartShellProps.dir` | Explicit LTR/RTL overrides ancestor Fluent direction and SPFx culture direction. |
| `WebPartShellProps.action` | Header action content slot under both providers; shorthand retains its DOM. |
| `WebPartShellProps.children` | Application content; its mounted identity survives title mode and visibility changes. |
| `WebPartShellProps.dangerouslySetInnerHTML` | Disallowed; untyped injection is discarded. |

Native div attributes/events, `className`, consumer-provided `style` and the native div ref are forwarded to the shell layout root. The native `title` tooltip attribute is omitted from the props contract because title content belongs to persisted properties. There are no public editor options, engine types, title variants or new SPFx base classes.

### getWebPartShellPropertiesMetadata

Every call returns a fresh map and nested descriptor: `{ title: { isSearchablePlainText: true } }`. Compose it additively with the web part's own metadata as in the example. The helper has no Property Pane runtime dependency and does not mutate consumer metadata. `titleHTML` remains ordinary data; the helper does not classify it as `isHtmlString`. Search indexing and host save/reload still require authenticated SharePoint validation.

## Surface, theme and lifecycle

Auto surface selects application when the SPFx Microsoft 365 SDK is present or the guarded optional `legacyPageContext.pageLayoutType` equals exactly `SingleWebPartAppPage`. SDK presence is enough before asynchronous context acquisition completes. Missing, malformed or unknown legacy metadata falls back to section; the legacy field is not a stable detection guarantee. Use `surface="application"` when an application host lacks that signal. Surface detection uses neither viewport width nor manifest/form-factor inference, REST or polling.

Section adds zero automatic padding and a transparent background, including the Fluent wrapper. Application uses the existing medium spacing (12px) below 640px container inline-size and large spacing (16px) from 640px. Explicit padding wins. The container query measures the shell root, separately from its padded content; it is not a viewport breakpoint. Layout uses natural height and a wrapping title/action header. The shell adds no scroll lock, sticky header or SharePoint canvas sizing.

Header and body are separated by `gap.medium` (12px with the default Fluent theme) in every surface, including SharePoint sections. This internal gap is independent of outer padding. Hiding the title retains the gap when header actions are present; hiding it without actions removes the empty header and its gap. Content remains mounted when title visibility changes.

The title and header actions are vertically centered within their wrapping row using `alignItems.center`.

Theme tokens merge host, then the nearest ancestor Fluent theme, then the explicit theme override. Missing partial ancestor tokens come from the host; without a host theme, the hook supplies a light fallback. Direction resolves explicit `dir`, then an actual ancestor FluentProvider, then SPFx `cultureInfo.isRightToLeft`. A low-level Fluent context without the provider theme-class context does not establish ancestor direction. The shell supplies `targetDocument` from `instance.context.domElement.ownerDocument` and inherits Griffel's renderer. The host must supply a renderer for that document, including iframe hosts; the shell creates no renderer.

Same-instance rerenders preserve application content and actions. A different `instanceId` resets the keyed runtime subtree. Hiding or emptying the title leaves actions and children mounted. Title Read/Edit changes mount/destroy only the editor. The existing SPFx provider owns service readiness, instance isolation and subscriptions; dispose the React root in the web part's `onDispose`.

## Title authority and HTML protocol

With undefined `titleHTML`, `title` is literal text rendered as H2. A defined empty `titleHTML` means an empty title even when the old plain title is nonempty. Empty/whitespace-only documents render no heading in Read; Edit shows an unsaved placeholder. Recoverable imported HTML is normalized in memory without property writes. Unrecoverable input uses the literal plain fallback and an Edit-mode warning. Mounting, theme changes, focus and selection do not rewrite properties.

The private document contains one H1–H4 or Normal paragraph block, with text and hard breaks. Enter and Shift+Enter insert a hard break in that block outside IME composition. Multiblock paste joins block boundaries with hard breaks and uses the first supported block's format. These are V1 choices rather than a promise of exact native SharePoint editor parity.

Canonical HTML uses `p`, `h1`–`h4`, `strong`, `em`, `u`, `span` and `br`, escaped text, stable semantic classes and deterministic nesting (`strong` → `em` → `u` → `span`). Read renders validated React nodes and does not inject original HTML or load Tiptap.

| Formatting | Values and persisted markers |
| --- | --- |
| Font size | Integer 8–96px; `wps-size-N`. Suggestions: 8/10/12/14/16/18/20/24/28/32/36/48/60/72/96. |
| Block defaults | H1 → `title1` (32/40px), H2 → `title2` (28/36px), H3 → `title3` (24/32px), H4 → `subtitle1` (20/28px), Normal → `body1` (14/20px), without explicit overrides. Values shown are font size/line-height in the standard web theme; active theme tokens determine actual values. |
| Theme foreground | Stable `wps-color-<id>` markers from the closed foreground registry below. Automatic has no color mark. |
| Line spacing | Default uses the block's Fluent line-height token. Explicit 1/1.2/1.4/1.6/2/2.4 overrides use `wps-line-10/12/14/16/20/24`. “Default line spacing” removes only that override. |
| Space before/after | `wps-space-before` and `wps-space-after` toggles; active space uses `spacingVerticalS`. |
| Emphasis | Bold, italic and underline marks; no links, lists, images or arbitrary color/style configuration. |

Import accepts `b`/`i`, the observed legacy `fontSizeXLarge` alias as 20px, and `fontColorBlue` as the semantic brand role (no identical RGB guarantee). A supported unitless native `line-height` on a supported block can be imported; other inline styles, URLs, event handlers and unknown classes are discarded. Harmless wrappers flatten to text; active/embedded/form/media subtrees are removed. Persisted classes are protocol markers under scoped Griffel rules, not Griffel hashes. Theme changes alter tokens without rewriting HTML.

Read and Edit share the official Fluent typography recipes, including font family and weight. The font-size picker reflects inherited theme sizes; explicit `wps-size-N` and `wps-line-N` values remain authoritative. Existing titles without overrides adopt these presets without a property write or HTML migration; imported native `line-height:1.4` remains an explicit override. Semantic heading levels and visual presets are associated by this control, not automatically by Fluent.

Input and canonical output are limited to 65,536 UTF-8 bytes; parser nesting is limited to depth 32. Input that exceeds a limit or would exceed the canonical byte limit uses the safe fallback. An oversized editor draft stays visible with a recoverable size error; neither title property is patched until the draft fits. There is no implicit truncation. Leaving Edit discards an unsaved oversized draft when the editor is destroyed; previously saved properties remain authoritative.

## Theme colors and localized UI

The circular Fluent SwatchPicker contains vertically stacked Automatic, Theme colors and Standard colors sections, each with a horizontal wrapping row of swatches. Every swatch has a localized name and Tooltip available on pointer hover and keyboard focus; the accessible name matches that name. Theme colors describe roles rather than promising a fixed black/blue value. The same picker and names are available in the toolbar overflow.

The default light-theme palette follows the reference order: 12 Theme colors and 10 Standard colors. The values come exclusively from foreground tokens; fixed `color-mix(in srgb, …)` recipes provide intermediate tones. Their appearance follows the current Fluent theme and may differ from the older Office reference palette, particularly its cyan, blue and purple.

| Offered Theme ID | Foreground source |
| --- | --- |
| `primary` | `colorNeutralForeground1` |
| `neutralDark` | 50% neutral foreground 1 + 50% neutral foreground 2 |
| `subtle` | `colorNeutralForeground2` |
| `neutralMuted` | `colorNeutralForeground3` |
| `neutralLight` | 75% neutral foreground 3 + 25% neutral foreground inverted |
| `neutralLighter` | 60% neutral foreground 3 + 40% neutral foreground inverted |
| `brandDarkest` | 25% brand foreground 1 + 75% neutral foreground 1 |
| `brandDarker` | 40% brand foreground 1 + 60% neutral foreground 1 |
| `brandDark` | 55% brand foreground 1 + 45% neutral foreground 1 |
| `brandAlt` | `colorBrandForeground2` |
| `brand` | `colorBrandForeground1` |
| `brandLight` | 75% brand foreground 1 + 25% neutral foreground inverted |

| Offered Standard ID | Foreground source |
| --- | --- |
| `darkRed` | `colorPaletteRedForeground1` |
| `red` | `colorPaletteRedForeground3` |
| `amber` | `colorPaletteMarigoldForeground3` |
| `yellow` | `colorPaletteYellowForeground3` |
| `yellowGreen` | 73% yellow foreground 3 + 27% green foreground 3 |
| `green` | `colorPaletteGreenForeground3` |
| `cyan` | `colorPaletteLightTealForeground2` |
| `blue` | `colorPaletteBlueForeground2` |
| `navy` | `colorPaletteNavyForeground2` |
| `purple` | `colorPalettePurpleForeground2` |

The accepted protocol retains four earlier colors: `orange` (dark-orange foreground 3), `teal`, `pink` and `brown` (their palette foreground 2). These remain valid in saved titles and appear in the picker while selected, without expanding the default palette. All 26 IDs use stable `wps-color-<id>` markers; derived tones never save CSS expressions or resolved colors in the document.

Colors resolve from the shell's effective Fluent theme, including host/ancestor/explicit overrides. Duplicate resolved colors are collapsed while preserving the currently selected identity and Automatic. For example, default dark-theme red and dark red coincide, so only one is offered at a time; both protocol IDs remain valid. Automatic removes only the color override, preserving emphasis, size and block formatting. Existing primary/subtle/brand titles remain valid. No custom HEX/RGB color picker is provided.

```tsx
// The saved semantic green follows the active theme; its UI name is translated.
// The properties object belongs to the SPFx web part instance passed to WebPartShell.
const properties: IWebPartShellProperties = {
  title: 'Project documents',
  titleHTML: '<h2><span class="wps-color-green">Project documents</span></h2>',
  showTitle: true,
};
```

The control ships complete private catalogs for the [50 SharePoint UI cultures listed by Microsoft](https://support.microsoft.com/en-au/sharepoint/platform/languages-supported-by-sharepoint), recorded in the versioned [culture fixture](../../../tests/fixtures/sharepoint-ui-locales.json). The runtime reads `instance.context.pageContext.cultureInfo.currentUICultureName`, independently of the regional `currentCultureName`. Catalog selection normalizes case/separators, uses exact culture and explicit supported aliases, and falls back to `en-US` for absent or unsupported values. Portuguese regional variants, Chinese variants and Serbian scripts remain distinct. No extra locale prop, SPFx AMD string resource configuration or translation service is required.

Localization covers the editor and floating toolbar, every color name, overflow menus, heading/size/spacing choices, mixed-state and shortcut descriptions, empty-title placeholder, loading label, Retry, invalid-saved-title warning and oversized-draft error. Labels belong to each shell instance. Language updates preserve the existing editor view, author content, stored semantic color IDs, selection and undo history without publishing properties. Direction follows the existing `dir`/ancestor/SPFx contract separately from the UI language.

Translation key/placeholder checks and browser tests establish resource coverage and UI binding; they do not certify native-speaker linguistic quality. Authenticated SharePoint and screen-reader checks remain separate from local browser results.

## Editing and recovery

All author UI uses Fluent UI 9: floating Popover/Toolbar, heading Dropdown, validated font-size Combobox, finite theme SwatchPicker, spacing Menu, clear formatting, Overflow commands, Spinner and MessageBar. Tiptap manages only the contenteditable engine. All control-owned labels, tooltips, accessible names, placeholders and recovery messages follow the instance’s SharePoint UI culture. Browser spellcheck remains available; author text and consumer-owned content are never translated.

With a selected text range, inline formatting applies to that range. Without a text selection, bold, italic, underline, font size, color and Clear apply to the entire title while preserving the caret position. Ctrl/Cmd+B/I/U follows the same scope. Mixed emphasis becomes uniform on the first action and is removed on the next action; toolbar state describes the whole title at a caret. Color and size changes preserve each text run's unrelated overrides. An empty title stores formatting for subsequent typing without publishing empty-title properties.

Each formatting gesture forms one undo boundary. Clear removes targeted inline marks and block spacing overrides while keeping text and heading level. Heading and spacing already address the title's single block. Popup commands preserve the original caret/range and return editor focus where appropriate; keyboard Tab remains available. Icon buttons and toggles use the same medium size; heading and font-size selectors use small size and are vertically centered with compact widths. The toolbar fits its control content, keeps a theme-token gap between heading and font size, and moves commands into Fluent Overflow when space is limited. Format painter is intentionally absent in V1. Escape does not revert content already published.

One editor view exists per Edit session. Effective document transactions publish one merged `{ title, titleHTML }` patch, preserving other application properties; plain search text comes from the same validated document with breaks/whitespace normalized. Canonical echoes preserve selection/history. A real external title replacement updates the existing view, clamps its selection and establishes fresh history so Undo cannot restore the preceding external version. Theme/direction changes leave history intact.

During IME composition, external replacements wait. The latest pending replacement applies after composition only if a newer local edit has not superseded it; local composition wins over older external updates. Effective changes publish live, with a final flush on editor disposal; there is no save debounce. Local simulated composition cannot establish every device IME's behavior.

The visible Edit title lazily requests the editor. Loading keeps the read title with a Spinner and makes no writes. A chunk/engine failure is contained inside the title boundary and offers Retry, leaving application content mounted. Retry creates a fresh lazy attempt; it cannot guarantee recovery from an unavailable asset, cache/network failure or repeating engine error. Invalid saved input warnings and oversized draft errors are separate from load errors; application rendering errors are not hidden by a shell-wide boundary.

## Dependencies and package cost

Use the existing compatible shared React, SPFx, Fluent UI 9 and Griffel peers. The verified repository baseline is React/ReactDOM 17.0.1, TypeScript 5.3.3, SPFx 1.21.1 and Fluent 9.74.9; peer ranges are not a tested version matrix. Fourteen private implementation dependencies are pinned to Tiptap 3.30.3: core, react, pm, extensions, and extension-bold/bubble-menu/document/floating-menu/hard-break/heading/italic/paragraph/text/underline. Consumers need no new editor peer or public engine types. The optional React binding menu packages are pinned for compatible installation but are not imported by the title UI.

The locked added dependency subtree contains 32 package records, 7,573,815 raw installed bytes (about 7.6 MB), and 880 files. This is installed uncompressed content, not download size or a fresh-install guarantee. Installation footprint and emitted assets are separate measurements.

Root, `/controls` and `/controls/web-part-shell` share canonical component/type/helper identity. In the production contract, unrelated consumers exclude shell/editor/engine; the shell initial chunks exclude Tiptap/ProseMirror, and visible Edit loads the async editor. Metadata-only consumers exclude the shell runtime/providers/editor. Optional Bubble/Floating menu modules are excluded from emitted assets. The async `@tiptap/extensions` root uses Placeholder/UndoRedo but retains six other root extension definitions because of upstream factory effects; full extension pruning is not promised. Absolute asset sizes depend on the final source and toolchain; measure initial, async and union raw/gzip/Brotli totals after each relevant build instead of treating an earlier snapshot as current. See [package imports](../../PACKAGE-IMPORTS.md#webpartshell-imports-and-editor-boundary).

## Sample and verification limits

The actual sample root uses the shell, manifest defaults, searchable metadata and a show-title pane toggle. Its lazy **WebPartShell** panel makes surfaces/padding, inherited host theme/direction, property replacement, visibility and preserved child/action state observable. Dedicated browser fixtures cover native ref, event and style forwarding. Build the library before the app and restart an active serve after library changes. See [development checks](../../DEVELOPMENT.md#webpartshell-development) and the [authenticated checklist](../../SHAREPOINT-VALIDATION.md#webpartshell).

Local Chrome acceptance covers both source and compiled-library fixtures. `scripts/check-title-toolbar-localization-browser.cjs` checks the 50 UI catalogs, color tooltips, instance isolation and narrow/RTL layouts; `scripts/check-title-toolbar-geometry-browser.cjs` checks foreground rendering, whole-title formatting, overflow, themes, forced colors and native 200% zoom. These fixtures mount real providers/hooks with service doubles; composition events use CDP and paste uses a synthetic ClipboardEvent, alongside actual mouse/keyboard interactions. They establish local behavior, not authenticated host save/reload/search, live Teams/Office/Outlook behavior, device IME, touch or screen-reader acceptance. Those checks remain NOT EXECUTED until their own host observations are recorded. Catalog completeness and localized UI bindings do not establish native-speaker certification of the translations.

## Header action slot

`action` accepts string/JSX shorthand or a fixed `div` props object under the existing SPFx and Fluent providers. Shorthand keeps its existing DOM; object refs target an inner action div, separate from the shell root. Native attributes, className, consumer style and Fluent render callbacks are supported with the existing peer dependencies:

```tsx
<WebPartShell instance={this} action={{
  className: classes.actions,
  children: (Component, props) => <Component {...props}><MyActions /></Component>,
}}><MyContent /></WebPartShell>
```

An empty props object requests the header/action container even when the persisted title is hidden. Booleans, undefined and legacy null suppress action content. Callbacks execute only during ordinary rendering, may run more than once, and are never invoked for presence inspection. Static strings are literal. `as` can only repeat div; raw HTML and unsupported native keys are discarded. JSX/callback subtrees remain caller-owned, and React 17 erased `ReactNode` types can bypass literal-object restrictions. The persisted title protocol, title editor, provider ownership and body lifecycle remain unchanged.

[View public source](../../../packages/spfx-react-toolkit/src/controls/web-part-shell/index.ts)
