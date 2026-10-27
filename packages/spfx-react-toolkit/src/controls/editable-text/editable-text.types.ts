import type * as React from 'react';

/**
 * Semantic block tags supported by EditableText's single formatted region.
 * Use these values in `allowedBlocks` and `defaultBlock`; H5/H6 are unsupported.
 * Omitted `allowedBlocks` admits every value. Omitted `defaultBlock` selects `p`
 * when allowed, otherwise the lowest allowed heading level.
 *
 * @example
 * import type { EditableTextBlock } from '@apvee/spfx-react-toolkit';
 * const headings: readonly EditableTextBlock[] = ['h2', 'h3'];
 */
export type EditableTextBlock = 'p' | 'h1' | 'h2' | 'h3' | 'h4';

/**
 * Controlled value emitted after an effective EditableText user change.
 * The caller reflects `html` into the controlled value and owns persistence.
 * Configuration normalization does not emit a change. Deletion emits two empty strings.
 *
 * @example
 * import type { EditableTextChangeData } from '@apvee/spfx-react-toolkit';
 * const change: EditableTextChangeData = { html: '<p>News<br>Today</p>', text: 'News Today' };
 */
export interface EditableTextChangeData {
  /** Canonical escaped toolkit HTML, or `''` for an empty region; hard breaks use `br`. */
  readonly html: string;
  /** Searchable plain text with whitespace and hard breaks collapsed to spaces and trimmed. */
  readonly text: string;
}

/**
 * Accessible swatch names keyed by the editor's stable semantic colors. These
 * labels never change persisted color identifiers or theme-derived values.
 * Every member is required in a complete dictionary; use Partial for overrides.
 * @example
 * const customizeColors = (names: EditableTextColorNames): EditableTextColorNames =>
 *   ({ ...names, red: 'Danger', brand: 'Accent' });
 */
export interface EditableTextColorNames {
  /** Neutral tone 2 semantic color tooltip and accessible swatch name. */
  readonly neutralDark: string;
  /** Neutral tone 5 semantic color tooltip and accessible swatch name. */
  readonly neutralLight: string;
  /** Neutral tone 6 semantic color tooltip and accessible swatch name. */
  readonly neutralLighter: string;
  /** Brand tone 1 semantic color tooltip and accessible swatch name. */
  readonly brandDarkest: string;
  /** Brand tone 2 semantic color tooltip and accessible swatch name. */
  readonly brandDarker: string;
  /** Brand tone 3 semantic color tooltip and accessible swatch name. */
  readonly brandDark: string;
  /** Brand tone 6 semantic color tooltip and accessible swatch name. */
  readonly brandLight: string;
  /** Dark red semantic color tooltip and accessible swatch name. */
  readonly darkRed: string;
  /** Amber semantic color tooltip and accessible swatch name. */
  readonly amber: string;
  /** Yellow green semantic color tooltip and accessible swatch name. */
  readonly yellowGreen: string;
  /** Cyan semantic color tooltip and accessible swatch name. */
  readonly cyan: string;
  /** Navy blue semantic color tooltip and accessible swatch name. */
  readonly navy: string;
  /** Primary text color semantic color tooltip and accessible swatch name. */
  readonly primary: string;
  /** Secondary text color semantic color tooltip and accessible swatch name. */
  readonly subtle: string;
  /** Brand color semantic color tooltip and accessible swatch name. */
  readonly brand: string;
  /** Muted text color semantic color tooltip and accessible swatch name. */
  readonly neutralMuted: string;
  /** Alternate brand color semantic color tooltip and accessible swatch name. */
  readonly brandAlt: string;
  /** Red semantic color tooltip and accessible swatch name. */
  readonly red: string;
  /** Orange semantic color tooltip and accessible swatch name. */
  readonly orange: string;
  /** Yellow semantic color tooltip and accessible swatch name. */
  readonly yellow: string;
  /** Green semantic color tooltip and accessible swatch name. */
  readonly green: string;
  /** Teal semantic color tooltip and accessible swatch name. */
  readonly teal: string;
  /** Blue semantic color tooltip and accessible swatch name. */
  readonly blue: string;
  /** Purple semantic color tooltip and accessible swatch name. */
  readonly purple: string;
  /** Pink semantic color tooltip and accessible swatch name. */
  readonly pink: string;
  /** Brown semantic color tooltip and accessible swatch name. */
  readonly brown: string;
}

/**
 * Complete observable editor UI dictionary. Omitted prop fields resolve from the
 * host-selected native SPFx resource. Labels do not translate saved author text,
 * configure direction or replace the caller's required textbox name. Formatter
 * callbacks run during rendering and must return strings without side effects.
 * @example
 * const customize = (labels: EditableTextStrings): EditableTextStrings => ({
 *   ...labels, bold: 'Strong', headingLevel: level => `Heading ${level}`,
 *   colorNames: { ...labels.colorNames, red: 'Danger' },
 * });
 */
export interface EditableTextStrings {
  /** Error shown when the lazy editor chunk fails to load. */
  readonly editorLoadError: string;
  /** Button label that retries the failed editor chunk. */
  readonly retry: string;
  /** Warning shown when saved markup falls back to literal text. */
  readonly savedFormattingWarning: string;
  /** Spinner label while the editor chunk loads. */
  readonly loadingEditor: string;
  /** Empty editor hint when the direct placeholder prop is omitted; never an accessible name. */
  readonly placeholder: string;
  /** Error shown when the live draft exceeds the serialized size limit. */
  readonly tooLargeError: string;
  /** Accessible name and tooltip for overflow formatting commands. */
  readonly moreFormatting: string;
  /** Heading selector and overflow submenu label. */
  readonly heading: string;
  /** Paragraph option label. */
  readonly normal: string;
  /** Bold command label and accessible name. */
  readonly bold: string;
  /** Italic command label and accessible name. */
  readonly italic: string;
  /** Underline command label and accessible name. */
  readonly underline: string;
  /** Color command label and accessible name. */
  readonly themeColor: string;
  /** Clear-formatting command label. */
  readonly clearFormatting: string;
  /** Spacing command label. */
  readonly spacing: string;
  /** Default line spacing option label. */
  readonly defaultLineSpacing: string;
  /** Space-before toggle label. */
  readonly spaceBefore: string;
  /** Space-after toggle label. */
  readonly spaceAfter: string;
  /** Accessible name of the formatting popup. */
  readonly formattingControls: string;
  /** Accessible name of its toolbar. */
  readonly formatting: string;
  /** Font size combobox and expand-button accessible name. */
  readonly fontSize: string;
  /** Mixed font-size placeholder. */
  readonly mixed: string;
  /** Accessible name of the color picker. */
  readonly themeColorGroup: string;
  /** Automatic color option and section label. */
  readonly automatic: string;
  /** Theme swatch section label. */
  readonly themeColors: string;
  /** Standard swatch section label. */
  readonly paletteColors: string;
  /** Accessible swatch labels; complete dictionaries require every semantic color. */
  readonly colorNames: EditableTextColorNames;
  /** Formats an offered heading level (1–4) into its selector label. */
  readonly headingLevel: (level: number) => string;
  /** Formats an offered numeric line-spacing multiplier into its menu label. */
  readonly lineSpacingValue: (value: number) => string;
  /** Combines a translated command label and literal keyboard shortcut for a tooltip. */
  readonly formattingShortcut: (label: string, shortcut: string) => string;
  /** Describes a command or group with mixed selection formatting using its label. */
  readonly mixedFormatting: (label: string) => string;
}

/**
 * Partial instance dictionary with independently optional semantic color names.
 * Omitted or undefined fields inherit native SPFx labels; empty strings remain
 * intentional overrides. Direct placeholder and ARIA props retain precedence.
 * Replacing this dictionary preserves the editor, draft, selection and undo
 * history and never invokes onChange. No provider or translation service is needed.
 * @example
 * const labels: EditableTextStringsOverride = {
 *   placeholder: 'Write a summary', colorNames: { brand: 'Accent' },
 *   formattingShortcut: (label, shortcut) => `${label}: ${shortcut}`,
 * };
 */
export type EditableTextStringsOverride = Partial<Omit<EditableTextStrings, 'colorNames'>> & {
  /** Overrides individual semantic swatch names; omitted colors retain native labels. */
  readonly colorNames?: Partial<EditableTextColorNames>;
};

type EditableTextNativeProps = Omit<React.HTMLAttributes<HTMLDivElement>,
  'children' | 'dangerouslySetInnerHTML' | 'contentEditable' | 'suppressContentEditableWarning' |
  'defaultValue' | 'role' | 'dir' | 'onChange' | 'placeholder' | 'aria-label' | 'aria-labelledby' | 'aria-describedby'>;

/**
 * Controlled single-block formatted text, independent of SPFx providers.
 * Native React div attributes, events, `className` and consumer `style` target the
 * stable root. The forwarded ref also targets that div in both modes. Content,
 * role, contentEditable, defaultValue and direction are owned by the control;
 * configure theme, direction and popup document on the host FluentProvider and
 * Griffel RendererProvider. The control creates neither provider nor renderer.
 *
 * Edit requires a callback and at least one accessible-name attribute. A dynamic
 * Read/Edit mode is supported when both requirements are supplied. Read defaults
 * to no callback. Type checking cannot prove that a string is nonempty: Edit
 * validates string inputs, the name and callback at runtime. Invalid inputs or block configuration throw
 * TypeError in either mode. This browser control owns p/h1–h4 nodes and must be
 * placed in a container that permits block elements, outside labels or triggers.
 * No editor/AST types, toolbar slots or persistence integration are exposed.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { EditableText, type EditableTextProps } from '@apvee/spfx-react-toolkit';
 * function Summary(): React.ReactElement {
 *   const [html, setHtml] = React.useState<string>();
 *   const props: EditableTextProps = {
 *     mode: 'edit', value: html, fallbackText: 'Summary',
 *     allowedBlocks: ['h2', 'h4'], 'aria-label': 'Summary',
 *     onChange: data => setHtml(data.html),
 *   };
 *   return <EditableText {...props} />;
 * }
 * ```
 */
export type EditableTextProps = EditableTextNativeProps & {
  /**
   * Controlled toolkit HTML. Omission uses literal fallbackText; `''` is an
   * authoritative empty value. Unrecoverable HTML uses fallbackText and shows an
   * Edit warning. Safe imports and block conversion normalize only in memory,
   * without onChange. The caller must reflect user changes here and owns saving.
   * Canonical echoes preserve selection/history; actual source replacements
   * establish a new undo baseline. Input/output is limited to 65,536 UTF-8 bytes
   * and parsing depth 32; unsupported markup is never injected as raw HTML.
   */
  value?: string;
  /**
   * Literal fallback when value is omitted or unrecoverable; defaults to `''`.
   * Never interpreted as HTML or synchronized as a second controlled value.
   * Changing value/fallbackText is an external source update, unlike configuration.
   */
  fallbackText?: string;
  /**
   * Semantic block membership in Read and every Edit input path. Omission admits
   * h1/h2/h3/h4/p. Order and duplicates are ignored; an empty array or unsupported
   * runtime member throws TypeError. A membership change reconciles the live
   * draft and resets undo while retaining the editor, text, marks and valid
   * selection. It never publishes configuration changes. A singleton hides the
   * block selector. Font size, emphasis, color and spacing remain independent.
   */
  allowedBlocks?: readonly EditableTextBlock[];
  /**
   * Initial/fallback/empty block and conversion target for excluded blocks.
   * Defaults to p if allowed, otherwise the lowest allowed heading. An explicit
   * excluded value throws TypeError. Changing only this default preserves undo
   * and any nonempty permitted block, without publishing. During IME the latest
   * policy applies after composition; configuration alone never produces changes.
   */
  defaultBlock?: EditableTextBlock;
  /**
   * Literal empty-editor hint, never serialized or used as its accessible name.
   * Omission uses the generic localized hint; `''` suppresses the hint. Has no
   * effect in Read. Changes preserve draft, selection and history.
   */
  placeholder?: string;
  /**
   * Instance UI labels and pure formatting callbacks over native SPFx resources.
   * Omitted/undefined fields use native labels; colorNames merges by semantic
   * field. Empty strings remain intentional. Direct placeholder wins over the
   * dictionary hint. Changes preserve draft, selection and history without
   * publishing onChange; saved author content and host direction are unchanged.
   */
  strings?: EditableTextStringsOverride;
  /**
   * Accessible name on the actual Edit textbox, or on the Read root. Edit
   * requires this or aria-labelledby to contain non-whitespace text. Omission
   * supplies no implicit name; placeholder and dictionaries do not replace it.
   */
  'aria-label'?: string;
  /**
   * Space-separated IDs naming the Edit textbox, or the Read root. At least
   * this or aria-label must be nonempty in Edit. The caller owns the referenced
   * labeling elements; omission creates no association.
   */
  'aria-labelledby'?: string;
  /**
   * Optional space-separated description IDs for the actual Edit textbox or
   * Read root; the caller owns these elements. Omission creates no association.
   */
  'aria-describedby'?: string;
} & (
  {
    /** Rendering mode, default read. Only edit loads the private editor engine. */
    mode?: 'read';
    /**
     * Optional domain callback in Read; Read originates no changes. An Edit→Read
     * transition finishes pending Edit input using the latest provided callback,
     * or the last committed Edit callback if the Read callback is omitted. In
     * Edit it receives live effective changes plus the final blur/composition/
     * unmount flush, with the latest committed block policy on mode exit.
     * Receives canonical HTML and normalized search text from one validated
     * document. Deletion emits `{ html: '', text: '' }`; formatting an empty
     * document prepares typing without emitting. Mount, external value and
     * policy normalization never emit. Oversized live drafts remain visible
     * with an error and suppress publication instead of emitting partial data.
     */
    onChange?: (data: EditableTextChangeData) => void;
  } | ({
    /** Explicit edit, or a dynamic read/edit mode with callback and name supplied. */
    mode: 'read' | 'edit';
    /**
     * Required controlled domain callback in Edit. Reflect data.html into value;
     * data.text is whitespace-normalized search text. Emission, final flush and
     * size-error behavior follow the Read branch's onChange contract above.
     */
    onChange: (data: EditableTextChangeData) => void;
  } & ({
    /** Required nonempty Edit textbox name when aria-labelledby is omitted; see the common name contract. */
    'aria-label': string;
  } | {
    /** Required nonempty Edit textbox label IDs when aria-label is omitted; see the common labeling contract. */
    'aria-labelledby': string;
  }))
);
