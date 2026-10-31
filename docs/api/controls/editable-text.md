# EditableText

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > EditableText

`EditableText` owns one formatted block with hard breaks. It is controlled React UI with no SPFx provider requirement. Native SPFx AMD resources supply its UI labels. Read mode renders safe content; Edit loads the private editor on demand. The caller owns the saved HTML and reflects user changes into `value`. Mounting, loading external content and changing configuration never save normalized HTML automatically.

Import `EditableText`, `EditableTextProps`, `EditableTextBlock`, `EditableTextChangeData`, `EditableTextStrings`, `EditableTextColorNames` and `EditableTextStringsOverride` from `@apvee/spfx-react-toolkit`, `/controls` or `/controls/editable-text`. These resolve to the same public implementation and types. There are no public engine, AST, codec, toolbar slot or plugin exports.

## React usage

Under the application's existing FluentProvider and document's Griffel RendererProvider:

```tsx
import * as React from 'react';
import {
  EditableText, type EditableTextProps,
  type EditableTextBlock, type EditableTextChangeData,
} from '@apvee/spfx-react-toolkit/controls/editable-text';

const headings: readonly EditableTextBlock[] = ['h2', 'h3', 'h4'];

export function ProjectSummary(): React.ReactElement {
  const [html, setHTML] = React.useState<string>();
  const [text, setText] = React.useState('Project summary');
  const [editing, setEditing] = React.useState(false);
  const root = React.useRef<HTMLDivElement>(null); // React 17 ref initializer
  const change = (data: EditableTextChangeData): void => {
    setHTML(data.html);
    setText(data.text);
  };
  const props: EditableTextProps = {
    mode: editing ? 'edit' : 'read', value: html,
    fallbackText: 'Project summary', allowedBlocks: headings,
    'aria-label': 'Project summary', 'aria-describedby': 'summary-help',
    onChange: change, strings: { bold: 'Strong', colorNames: { red: 'Danger' } },
  };
  return <section>
    <button type="button" onClick={() => setEditing(value => !value)}>Toggle editing</button>
    <p id="summary-help">One heading; Enter inserts a hard break.</p>
    <EditableText {...props} ref={root} data-project-summary="" />
    <output>{text}</output>
  </section>;
}
```

React 17's `useRef<HTMLDivElement>(null)` interoperates with the legacy React ref API; application absence values use `undefined`. No SPFx hook/provider is required. Share the host's existing Fluent/Griffel peers; the control creates no provider or renderer. The host sets theme, direction, popup target document and renderer target document, including iframe/secondary-document scenarios.

The stable native `div` owns a `p` or H1–H4 region. Place it in a block-capable container, outside labels, buttons, triggers and other inline-only slots. It is not a general replacement for arbitrary inline slot content. Use [FormattedText](formatted-text.md) for read-only Markdown/JSX rendering.

## Public contract

`EditableTextBlock` is `'p' | 'h1' | 'h2' | 'h3' | 'h4'`. `EditableTextProps` inherits the supported React native div attributes/events, `className` and consumer `style`; these and the forwarded `ref` target the same outer div in both modes. The control owns children, `dangerouslySetInnerHTML`, `contentEditable`, `suppressContentEditableWarning`, `defaultValue`, `role` and `dir`. Use FluentProvider direction rather than a control `dir` prop. Native `onChange` is replaced by the domain callback below.

| Member | Observable behavior |
| --- | --- |
| `EditableTextProps.value` | Optional controlled toolkit HTML. `undefined` uses literal fallback; `''` is authoritative empty content. Unrecoverable HTML uses fallback with an Edit warning. Safe import and block conversion happen in memory without a callback. |
| `EditableTextProps.fallbackText` | Optional literal text, default `''`. Never parsed as HTML and never a second synchronized value. Used only for omitted/unrecoverable `value`. |
| `EditableTextProps.mode` | Optional `'read'` by default. `'edit'` requires a callback and accessible name. A dynamic `'read' \| 'edit'` mode requires both as well. Read loads no editor; leaving Edit destroys that instance's editor session. |
| `EditableTextProps.onChange` | Required in Edit, optional in Read. Receives `EditableTextChangeData` after effective user changes and final pending blur/composition/unmount flush. Reflect `data.html` into `value`. Read itself originates no changes. No event parameter is supplied. |
| `EditableTextProps.allowedBlocks` | Optional readonly `EditableTextBlock[]`; omitted allows p and H1–H4. Membership controls every Read/Edit input path. Order and duplicate entries do not matter. A singleton hides the block selector. |
| `EditableTextProps.defaultBlock` | Optional initial/fallback/empty block and excluded-block conversion target. Defaults to p when admitted, otherwise the lowest admitted heading. Must belong to the allowed set. |
| `EditableTextProps.placeholder` | Optional literal empty-editor hint. Omission uses localized generic text; `''` suppresses it. Never saved or used as the accessible name; no Read effect. |
| `EditableTextProps.strings` | Optional `EditableTextStringsOverride`: instance labels and formatter callbacks over native SPFx resources. Nested color names merge individually; omitted/undefined fields inherit native labels and empty strings remain intentional. |
| `EditableTextProps.'aria-label'` | Optional in Read; in Edit this or `aria-labelledby` must contain non-whitespace text. Routed to the actual Edit textbox or Read root. No implicit name. |
| `EditableTextProps.'aria-labelledby'` | Optional space-separated IDs, with the same Edit requirement. Caller supplies referenced labeling elements. Omission creates no association. |
| `EditableTextProps.'aria-describedby'` | Optional description IDs on the actual Edit textbox or Read root. Caller supplies the description elements. Omission creates no association. |
| `EditableTextChangeData.html` | Canonical escaped HTML with hard breaks as `br`; effective deletion emits `''`. |
| `EditableTextChangeData.text` | Searchable text: whitespace and breaks collapse to spaces and are trimmed. Effective deletion emits `''`. |

Runtime validation throws `TypeError` for invalid strings, mode, callback, accessible Edit name or block policy. TypeScript cannot guarantee nonempty string values. An empty allowed array, unsupported runtime member (including H5/H6), or explicit default outside the set is invalid. Equivalent arrays such as `['h2', 'h4']` and `['h4', 'h2', 'h2']` represent the same membership.

For example, `<h2>Project</h2>` under `allowedBlocks={['p']}` renders a paragraph. Stored input remains the original H2 until a user edit publishes a paragraph. This is valid policy conversion, not an import error or fallback. Read and Edit initialization, paste, external replacement and reset use the same finite policy. Font size, emphasis, color and line spacing are independent of the semantic block choice.

## UI dictionaries and migration

`locale` has been removed in this major API migration. Configure the package's native SPFx localized resources in the consuming solution; SPFx selects one culture when loading its AMD module. Changing `pageContext` or a browser locale at runtime does not reload that module. The 50 canonical translator catalogs remain package build inputs. Outside the SPFx loader, the host must provide an explicit resource-module adapter even when supplying a complete custom dictionary.

Use `strings={{ bold: 'Strong', colorNames: { red: 'Danger' } }}` for per-instance overrides. `EditableTextStrings` is a complete dictionary and `EditableTextStringsOverride` is partial except for its independently partial `colorNames`. `EditableTextColorNames` has only public semantic fields and no private editor/AST dependency. `EditableTextStringsOverride.colorNames` overrides individual entries; its omission inherits all names. Complete dictionaries require every label, formatter and color; partial dictionaries inherit omitted/undefined fields. Empty strings are retained as explicit values. Direct `placeholder` takes precedence, including `''` to suppress the hint. Direct ARIA props alone name the textbox; dictionary hints cannot satisfy Edit's accessible-name requirement.

Formatter callbacks receive heading levels 1–4, numeric line spacing, a translated label plus literal shortcut, or a mixed-state label. They return UI strings, run during rendering and must be pure. Defaults use the selected native catalog's templates, preserving literal inserted labels and shortcuts. Replacing a dictionary updates toolbar/color/hint/loading/recovery labels on the existing editor without publishing `onChange`, changing author content, resetting selection or clearing undo history. Two instances can use different dictionaries under the same native culture. The sample dictionary selector demonstrates native labels, a complete Italian `EditableTextStrings` dictionary (all callbacks and semantic color names), and partial project-label overrides on the same subject. Theme and direction remain host concerns.

| Dictionary member | Observable behavior |
| --- | --- |
| `EditableTextStrings.editorLoadError` | Error shown when the lazy editor chunk fails to load. |
| `EditableTextStrings.retry` | Button label that retries the failed editor chunk. |
| `EditableTextStrings.savedFormattingWarning` | Warning shown when saved markup falls back to literal text. |
| `EditableTextStrings.loadingEditor` | Spinner label while the editor chunk loads. |
| `EditableTextStrings.placeholder` | Empty editor hint when the direct placeholder prop is omitted; never an accessible name. |
| `EditableTextStrings.tooLargeError` | Error shown when the live draft exceeds the serialized size limit. |
| `EditableTextStrings.moreFormatting` | Accessible name and tooltip for overflow formatting commands. |
| `EditableTextStrings.heading` | Heading selector and overflow submenu label. |
| `EditableTextStrings.normal` | Paragraph option label. |
| `EditableTextStrings.bold` | Bold command label and accessible name. |
| `EditableTextStrings.italic` | Italic command label and accessible name. |
| `EditableTextStrings.underline` | Underline command label and accessible name. |
| `EditableTextStrings.themeColor` | Color command label and accessible name. |
| `EditableTextStrings.clearFormatting` | Clear-formatting command label. |
| `EditableTextStrings.spacing` | Spacing command label. |
| `EditableTextStrings.defaultLineSpacing` | Default line spacing option label. |
| `EditableTextStrings.spaceBefore` | Space-before toggle label. |
| `EditableTextStrings.spaceAfter` | Space-after toggle label. |
| `EditableTextStrings.formattingControls` | Accessible name of the formatting popup. |
| `EditableTextStrings.formatting` | Accessible name of its toolbar. |
| `EditableTextStrings.fontSize` | Font size combobox and expand-button accessible name. |
| `EditableTextStrings.mixed` | Mixed font-size placeholder. |
| `EditableTextStrings.themeColorGroup` | Accessible name of the color picker. |
| `EditableTextStrings.automatic` | Automatic color option and section label. |
| `EditableTextStrings.themeColors` | Theme swatch section label. |
| `EditableTextStrings.paletteColors` | Standard swatch section label. |
| `EditableTextStrings.colorNames` | Accessible swatch labels; complete dictionaries require every semantic color. |
| `EditableTextStrings.headingLevel` | Formats an offered heading level (1–4) into its selector label. |
| `EditableTextStrings.lineSpacingValue` | Formats an offered numeric line-spacing multiplier into its menu label. |
| `EditableTextStrings.formattingShortcut` | Combines a translated command label and literal keyboard shortcut for a tooltip. |
| `EditableTextStrings.mixedFormatting` | Describes a command or group with mixed selection formatting using its label. |
| `EditableTextColorNames.neutralDark` | Neutral tone 2 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.neutralLight` | Neutral tone 5 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.neutralLighter` | Neutral tone 6 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brandDarkest` | Brand tone 1 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brandDarker` | Brand tone 2 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brandDark` | Brand tone 3 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brandLight` | Brand tone 6 semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.darkRed` | Dark red semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.amber` | Amber semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.yellowGreen` | Yellow green semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.cyan` | Cyan semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.navy` | Navy blue semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.primary` | Primary text color semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.subtle` | Secondary text color semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brand` | Brand color semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.neutralMuted` | Muted text color semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brandAlt` | Alternate brand color semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.red` | Red semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.orange` | Orange semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.yellow` | Yellow semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.green` | Green semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.teal` | Teal semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.blue` | Blue semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.purple` | Purple semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.pink` | Pink semantic color tooltip and accessible swatch name. |
| `EditableTextColorNames.brown` | Brown semantic color tooltip and accessible swatch name. |

## Color picker and theme colors

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

The accepted protocol retains four earlier colors: `orange` (dark-orange foreground 3), `teal`, `pink` and `brown` (their palette foreground 2). These remain valid in saved content and appear in the picker while selected, without expanding the default palette. All 26 IDs use stable `wps-color-<id>` markers; derived tones never save CSS expressions or resolved colors in the document.

Colors resolve from the active FluentProvider theme supplied by the host. Duplicate resolved colors are collapsed while preserving the currently selected identity and Automatic. For example, default dark-theme red and dark red coincide, so only one is offered at a time; both protocol IDs remain valid. Automatic removes only the color override, preserving emphasis, size and block formatting. Existing primary/subtle/brand content remains valid. No custom HEX/RGB color picker is provided.

```tsx
import * as React from 'react';
import { EditableText } from '@apvee/spfx-react-toolkit/controls/editable-text';

// The saved semantic green follows the active theme; its UI name is translated.
const savedHTML = '<h2><span class="wps-color-green">Project documents</span></h2>';

export function ProjectHeading(): React.ReactElement {
  return <EditableText value={savedHTML} />;
}
```

## Controlled updates, history and composition

Canonical echoes preserve the current selection and history. Actual external source replacement (`value`/`fallbackText`) establishes a new undo baseline. A membership change reconciles the live draft, retaining the editor, text, marks and valid selection while resetting undo, even if the current block is still admitted. Reordering/duplicates do not reset it. Changing only `defaultBlock` preserves history and a nonempty admitted block. Placeholder and dictionary updates preserve the draft, selection and history. None of these configuration operations emits a callback.

During IME composition the latest committed policy applies after composition. An Edit→Read transition finishes pending Edit input with the current policy and latest provided callback, including when all three change in the same render. If Read omits its optional callback, only that final pending Edit flush uses the last committed Edit callback. Read then originates no changes. A configuration-only exit publishes nothing. Effective deletion emits `{ html: '', text: '' }`; formatting an empty document prepares future typing without publishing empty markup.

## HTML and editing limits

Canonical content consists of one p/H1–H4 block, text, `br`, strong/emphasis/underline and the supported semantic formatting markers. Enter and Shift+Enter insert hard breaks; multiple incoming blocks normalize into the one region. Bold/italic/underline commands and Ctrl/Cmd+B/I/U retain selection and provide undo boundaries. A caret-only inline formatting command affects the whole nonempty region. Clear formatting retains the block type. The editor supports size 8–96 px, finite line spacing and theme-derived color choices; H5/H6, lists, links, images, Markdown conversion and multi-paragraph documents are outside the public editing model. Unsupported wrappers may contribute safe text; unsafe embedded/executable subtrees are discarded. Unsupported markup is never injected as raw HTML.

The persisted protocol is shared with [WebPartShell titles](web-part-shell.md#title-authority-and-html-protocol): existing `wps-*` semantic classes, escaped content, mark order and accepted legacy size/color imports are retained. Color IDs resolve against the current Fluent theme rather than storing arbitrary HEX colors. HTML input and output are limited to 65,536 UTF-8 bytes with parsing depth 32. Unrecoverable imports use literal fallback. An oversized live draft stays visible with an error and suppresses publication, without partial/truncated callbacks. Resolve the draft or replace its source before persisting it.

## Caller-owned SPFx persistence

Render this content under an initialized matching SPFx provider, for example [WebPartShell](web-part-shell.md). The existing hooks remain in the caller. They are not dependencies of EditableText itself.

```tsx
import * as React from 'react';
import {
  EditableText, useSPFxProperties, useSPFxDisplayMode,
  type IWebPartShellProperties,
} from '@apvee/spfx-react-toolkit';

interface Properties extends IWebPartShellProperties {
  summary: string;
  summaryHTML?: string;
}

export function SummaryContent(): React.ReactElement {
  const { properties, setProperties } = useSPFxProperties<Properties>();
  const { isEdit } = useSPFxDisplayMode();
  return <EditableText mode={isEdit ? 'edit' : 'read'}
    value={properties?.summaryHTML} fallbackText={properties?.summary ?? ''}
    allowedBlocks={['p']} aria-label="Project summary"
    onChange={data => setProperties({ summaryHTML: data.html, summary: data.text })} />;
}
```

`setProperties` shallow-merges partial property updates through the provider's host integration. This example preserves a plain search value beside canonical HTML. Empty HTML remains authoritative after reload. Property metadata, persistence permissions and any explicit saving workflow belong to the host application. Normalization alone writes neither property.

## Observable sample and verification

The lazy `editable-text` tab (`spfx-demo-editable-text`) consumes compiled library exports. It mounts one controlled subject under DemoCard, with all/paragraph/headings/H2–H4 presets, Read/Edit, H2/fallback/empty commands and separate raw stored input versus actual visible policy projection. The editor dictionary selector and last callback are in initially collapsed diagnostics; the selector changes per-instance labels rather than the native SPFx culture. Presets preserve the root; a Read transition ends only that editor session. The separate `tests/fixtures/demo-control-isolation/editable-text-regressions.tsx` accepts a real SPFx web part instance and places two independently controlled policies beside the Shell-owned title.

Automated DOM tests exercise the pure React panel and real control; compilation and local tests do not establish authenticated SharePoint behavior. See [SharePoint validation](../../development/sharepoint-validation/INDEX.md) for pending host, keyboard, IME and document-boundary checks.
