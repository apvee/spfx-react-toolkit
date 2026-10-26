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
   * UI culture for toolbar, hints and errors, with the toolkit's 50-culture
   * resolver/aliases and en-US fallback. Omission uses en-US. Does not translate
   * content or set direction; configure direction on FluentProvider. Live
   * locale changes retain editor, draft, selection and history.
   */
  locale?: string;
  /**
   * Accessible name on the actual Edit textbox, or on the Read root. Edit
   * requires this or aria-labelledby to contain non-whitespace text. Omission
   * supplies no implicit name; placeholder and locale do not replace it.
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
