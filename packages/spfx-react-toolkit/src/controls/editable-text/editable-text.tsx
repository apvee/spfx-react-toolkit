/**
 * Controlled single-block rich text with safe read rendering and lazy explicit editing.
 *
 * @packageDocumentation
 * @docPage controls/editable-text
 */
import * as React from 'react';
import { Button, MessageBar, MessageBarBody, Spinner } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import { getNativeElementProps } from '@fluentui/react-utilities';
import { getEditableTextPlainText, resolveEditableTextValue, serializeEditableTextDocument } from './editable-text-codec.internal';
import { resolveEditableTextBlockPolicy } from './editable-text-policy.internal';
import { EditableTextReadRenderer } from './editable-text-renderer.internal';
import { resolveEditableTextStrings } from './localization/editable-text-strings.internal';
import type { EditableTextStrings } from './editable-text.types';
import type { EditableTextDocument } from './editable-text-document.types.internal';
import type { EditableTextProps } from './editable-text.types';

function lazyEditor(): React.LazyExoticComponent<typeof import('./editable-text-editor.internal')['EditableTextEditor']> {
  return React.lazy(() => import(/* webpackChunkName: 'editable-text-editor' */ './editable-text-editor.internal')
    .then(module => ({ default: module.EditableTextEditor })));
}
// Module identity is shared; a lazy import starts only when rendered in Edit.
const InitialEditor = lazyEditor();

class EditableTextEditorBoundary extends React.Component<{
  children: React.ReactNode; onRetry: () => void; strings: EditableTextStrings;
}, { failed: boolean }> {
  public state = { failed: false };
  public static getDerivedStateFromError(): { failed: boolean } { return { failed: true }; }
  public render(): React.ReactNode {
    return this.state.failed ? <MessageBar intent="error"><MessageBarBody>{this.props.strings.editorLoadError}</MessageBarBody>
      <Button onClick={this.props.onRetry}>{this.props.strings.retry}</Button></MessageBar> : this.props.children;
  }
}

const ownedProps = ['value', 'fallbackText', 'mode', 'onChange', 'allowedBlocks', 'defaultBlock', 'placeholder', 'strings', 'locale',
  'children', 'dangerouslySetInnerHTML', 'contentEditable', 'suppressContentEditableWarning', 'defaultValue', 'role', 'dir', 'as'];

function validateProps(props: EditableTextProps): void {
  if (props.mode !== undefined && props.mode !== 'read' && props.mode !== 'edit') {
    throw new TypeError('EditableText mode must be read or edit');
  }
  for (const key of ['value', 'fallbackText', 'placeholder', 'aria-label', 'aria-labelledby', 'aria-describedby'] as const) {
    if (props[key] !== undefined && typeof props[key] !== 'string') throw new TypeError(`EditableText ${key} must be a string`);
  }
  if (props.onChange !== undefined && typeof props.onChange !== 'function') throw new TypeError('EditableText onChange must be a function');
  if (props.mode === 'edit') {
    if (typeof props.onChange !== 'function') throw new TypeError('EditableText Edit requires onChange');
    if (!props['aria-label']?.trim() && !props['aria-labelledby']?.trim()) {
      throw new TypeError('EditableText Edit requires a nonempty aria-label or aria-labelledby');
    }
  }
}

/**
 * A controlled, accessible region of single-block formatted text. Read renders
 * validated React nodes without loading the editor; Edit lazily loads the shared
 * private engine with per-instance draft/history and a local load-error Retry.
 * Normalization never persists data. Requires a browser document and the host's
 * Fluent/Griffel context for theme, direction and popups; no SPFx provider needed.
 * Configure native localized resource modules in SPFx; other hosts must supply
 * an explicit module adapter, including when every label is overridden.
 * Place the control in a block-capable container. See EditableTextProps for the
 * value protocol, controlled changes, limits and runtime TypeError conditions.
 *
 * @param props - Controlled value, mode, policy, accessible name and native div props.
 * @returns The stable div root in Read, loading, error and Edit states. A forwarded
 * HTMLDivElement ref targets this root; name/description ARIA targets the actual
 * textbox in Edit and the root in Read. Consumer style/events stay on the root.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { EditableText } from '@apvee/spfx-react-toolkit';
 * function Summary(): React.ReactElement {
 *   const [html, setHtml] = React.useState<string>();
 *   return <EditableText mode="edit" value={html} fallbackText="Summary"
 *     allowedBlocks={['p', 'h2']} aria-label="Summary"
 *     onChange={data => setHtml(data.html)} />;
 * }
 * ```
 *
 * @useWhen Bounded text needs safe read rendering with optional explicit rich-text editing.
 * @avoidWhen Use FormattedText for display-only inline Markdown without an editor lifecycle.
 *
 *
 * @example Control a project summary and preserve plain text
 * ```tsx
 * import * as React from 'react';
 * import {
 *   EditableText, type EditableTextProps,
 *   type EditableTextBlock, type EditableTextChangeData,
 * } from '@apvee/spfx-react-toolkit/controls/editable-text';
 *
 * const headings: readonly EditableTextBlock[] = ['h2', 'h3', 'h4'];
 *
 * export function ProjectSummary(): React.ReactElement {
 *   const [html, setHTML] = React.useState<string>();
 *   const [text, setText] = React.useState('Project summary');
 *   const [editing, setEditing] = React.useState(false);
 *   const root = React.useRef<HTMLDivElement>(null); // React 17 ref initializer
 *   const change = (data: EditableTextChangeData): void => {
 *     setHTML(data.html);
 *     setText(data.text);
 *   };
 *   const props: EditableTextProps = {
 *     mode: editing ? 'edit' : 'read', value: html,
 *     fallbackText: 'Project summary', allowedBlocks: headings,
 *     'aria-label': 'Project summary', 'aria-describedby': 'summary-help',
 *     onChange: change, strings: { bold: 'Strong', colorNames: { red: 'Danger' } },
 *   };
 *   return <section>
 *     <button type="button" onClick={() => setEditing(value => !value)}>Toggle editing</button>
 *     <p id="summary-help">One heading; Enter inserts a hard break.</p>
 *     <EditableText {...props} ref={root} data-project-summary="" />
 *     <output>{text}</output>
 *   </section>;
 * }
 * ```
 *
 *
 * @example Render saved semantic theme color
 * ```tsx
 * import * as React from 'react';
 * import { EditableText } from '@apvee/spfx-react-toolkit/controls/editable-text';
 *
 * // The saved semantic green follows the active theme; its UI name is translated.
 * const savedHTML = '<h2><span class="wps-color-green">Project documents</span></h2>';
 *
 * export function ProjectHeading(): React.ReactElement {
 *   return <EditableText value={savedHTML} />;
 * }
 * ```
 *
 *
 * @example Persist HTML and searchable text through SPFx properties
 * ```tsx
 * import * as React from 'react';
 * import {
 *   EditableText, useSPFxProperties, useSPFxDisplayMode,
 *   type IWebPartShellProperties,
 * } from '@apvee/spfx-react-toolkit';
 *
 * interface Properties extends IWebPartShellProperties {
 *   summary: string;
 *   summaryHTML?: string;
 * }
 *
 * export function SummaryContent(): React.ReactElement {
 *   const { properties, setProperties } = useSPFxProperties<Properties>();
 *   const { isEdit } = useSPFxDisplayMode();
 *   return <EditableText mode={isEdit ? 'edit' : 'read'}
 *     value={properties?.summaryHTML} fallbackText={properties?.summary ?? ''}
 *     allowedBlocks={['p']} aria-label="Project summary"
 *     onChange={data => setProperties({ summaryHTML: data.html, summary: data.text })} />;
 * }
 * ```
 *
 * @example Edit a controlled block with policy and dictionaries
 * ```tsx
 * // This standalone React scenario uses the host Fluent theme and Griffel renderer. Native resource module setup is required even with a complete dictionary. Read does not initialize this body editor; a surrounding Shell title can have its own session.
 * import * as React from 'react';
 * import { Button, Field, Input, Select } from '@fluentui/react-components';
 * import { EditableText, type EditableTextBlock, type EditableTextChangeData, type EditableTextStringsOverride, type EditableTextStrings } from '@apvee/spfx-react-toolkit';
 *
 * const italianEditableTextStrings: EditableTextStrings = {
 *   "editorLoadError": "Impossibile caricare l’editor del testo.",
 *   "retry": "Riprova",
 *   "savedFormattingWarning": "Impossibile leggere la formattazione salvata del testo. Modifica il testo per correggerla.",
 *   "loadingEditor": "Caricamento dell’editor del testo",
 *   "placeholder": "Aggiungi un testo",
 *   "tooLargeError": "Il testo è troppo grande per essere salvato. Accorcialo per continuare.",
 *   "moreFormatting": "Altre opzioni di formattazione del testo",
 *   "heading": "Intestazione",
 *   "normal": "Normale",
 *   "bold": "Grassetto",
 *   "italic": "Corsivo",
 *   "underline": "Sottolineato",
 *   "themeColor": "Colore del tema",
 *   "clearFormatting": "Cancella formattazione",
 *   "spacing": "Spaziatura",
 *   "defaultLineSpacing": "Interlinea predefinita",
 *   "spaceBefore": "Spazio prima",
 *   "spaceAfter": "Spazio dopo",
 *   "formattingControls": "Controlli di formattazione del testo",
 *   "formatting": "Formattazione del testo",
 *   "fontSize": "Dimensione carattere",
 *   "mixed": "Mista",
 *   "themeColorGroup": "Colori del tema del testo",
 *   "automatic": "Automatico",
 *   "themeColors": "Colori del tema",
 *   "paletteColors": "Colori standard",
 *   "colorNames": {
 *     "neutralDark": "Tonalità neutra 2",
 *     "neutralLight": "Tonalità neutra 5",
 *     "neutralLighter": "Tonalità neutra 6",
 *     "brandDarkest": "Tonalità del marchio 1",
 *     "brandDarker": "Tonalità del marchio 2",
 *     "brandDark": "Tonalità del marchio 3",
 *     "brandLight": "Tonalità del marchio 6",
 *     "darkRed": "Rosso scuro",
 *     "amber": "Ambra",
 *     "yellowGreen": "Verde giallastro",
 *     "cyan": "Ciano",
 *     "navy": "Blu navy",
 *     "primary": "Colore del testo principale",
 *     "subtle": "Colore del testo secondario",
 *     "brand": "Colore del marchio",
 *     "neutralMuted": "Colore del testo attenuato",
 *     "brandAlt": "Colore alternativo del marchio",
 *     "red": "Rosso",
 *     "orange": "Arancione",
 *     "yellow": "Giallo",
 *     "green": "Verde",
 *     "teal": "Verde petrolio",
 *     "blue": "Blu",
 *     "purple": "Viola",
 *     "pink": "Rosa",
 *     "brown": "Marrone"
 *   },
 *   headingLevel: level => `Intestazione ${level}`,
 *   lineSpacingValue: value => `Interlinea ${value}`,
 *   formattingShortcut: (label, shortcut) => `${label} (${shortcut})`,
 *   mixedFormatting: label => `${label}: formattazione mista`,
 * };
 *
 * const presets = ['all', 'paragraph', 'headings', 'h2-h4'] as const;
 * type Preset = typeof presets[number];
 * const blocks: Record<Preset, readonly EditableTextBlock[] | undefined> = {
 *   all: undefined,
 *   paragraph: ['p'],
 *   headings: ['h1', 'h2', 'h3', 'h4'],
 *   'h2-h4': ['h2', 'h3', 'h4'],
 * };
 * const customStrings: EditableTextStringsOverride = { bold: 'Strong', italic: 'Slanted', themeColor: 'Ink', formattingControls: 'Project text formatting', colorNames: { red: 'Alert red' } };
 * const initialHTML = '<h2><strong>Project</strong> update</h2>';
 *
 * // Pure React controlled example; SPFx persistence remains a caller concern.
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const [html, setHTML] = React.useState<string | undefined>(initialHTML);
 *   const [mode, setMode] = React.useState<'read' | 'edit'>('read');
 *   const [preset, setPreset] = React.useState<Preset>('all');
 *   const [fallback, setFallback] = React.useState('Literal <h2>fallback</h2>');
 *   const [placeholder, setPlaceholder] = React.useState('Write a project update');
 *   const [dictionary, setDictionary] = React.useState('native');
 *   const [lastChange, setLastChange] = React.useState<EditableTextChangeData>();
 *   const [changes, setChanges] = React.useState(0);
 *   // A reset invalidates the old session before any child cleanup. Only explicit Edit
 *   // admits a session into the current revision; Read still permits its ordinary final flush.
 *   const [editRevision, setEditRevision] = React.useState(resetRevision);
 *   const previousReset = React.useRef(resetRevision);
 *   const onChange = React.useCallback((data: EditableTextChangeData) => {
 *     if (editRevision !== resetRevision) return;
 *     setHTML(data.html);
 *     setLastChange(data);
 *     setChanges(count => count + 1);
 *   }, [editRevision, resetRevision]);
 *   React.useLayoutEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setHTML(initialHTML); setMode('read'); setPreset('all'); setFallback('Literal <h2>fallback</h2>');
 *     setPlaceholder('Write a project update'); setDictionary('native'); setLastChange(undefined); setChanges(0);
 *   }, [resetRevision]);
 *   return render({ controls: <>
 *       <Field label="Mode"><Select select={{ as: 'select', ...{ 'data-editable-text-demo': 'mode' } }} value={mode}
 *         onChange={event => {
 *           const next = event.target.value === 'edit' ? 'edit' : 'read';
 *           if (next === 'edit') setEditRevision(resetRevision);
 *           setMode(next);
 *         }}><option value="read">Read</option><option value="edit">Edit</option></Select></Field>
 *       <Field label="Allowed blocks"><Select select={{ as: 'select', ...{ 'data-editable-text-demo': 'preset', 'data-demo-variant': 'editable-text' } }} value={preset}
 *         onChange={event => { const selected = presets.find(value => value === event.target.value); if (selected !== undefined) setPreset(selected); }}>
 *         {presets.map(value => <option key={value} value={value}>{value}</option>)}
 *       </Select></Field>
 *       <Field label="Literal fallback"><Input input={{ as: 'input', ...{ 'data-editable-text-demo': 'fallback-input' } }} value={fallback} onChange={(_event, data) => setFallback(data.value)} /></Field>
 *       <Field label="Empty editor hint"><Input input={{ as: 'input', ...{ 'data-editable-text-demo': 'placeholder' } }} value={placeholder} onChange={(_event, data) => setPlaceholder(data.value)} /></Field>
 *     </>,
 *     commands: <>
 *       <Button data-editable-text-demo="h2" onClick={() => setHTML(initialHTML)}>Load H2 HTML</Button>
 *       <Button data-editable-text-demo="fallback" onClick={() => setHTML(undefined)}>Use literal fallback</Button>
 *       <Button data-editable-text-demo="empty" onClick={() => setHTML('')}>Load empty value</Button>
 *     </>,
 *     status: <>
 *       <p data-editable-text-demo="status" aria-live="polite">Mode: {mode} · Policy: {preset} · Callbacks: {changes}</p>
 *       <p>Stored controlled input: <code data-editable-text-demo="stored">{html === undefined ? 'undefined (literal fallback)' : html === '' ? '"" (authoritative empty)' : html}</code></p>
 *       <p>The live subject below is the visible policy projection. Normalization does not rewrite the stored input; a user change saves its canonical HTML into React state.</p>
 *     </>,
 *     diagnostics: <>
 *       <Field label="Editor dictionary"><Select select={{ as: 'select', ...{ 'data-editable-text-demo': 'strings' } }} value={dictionary} onChange={event => setDictionary(event.target.value)}>
 *         <option value="native">Native SPFx labels</option><option value="italian">Italiano · complete dictionary</option><option value="custom">Project labels · partial dictionary</option>
 *       </Select></Field>
 *       <p>Last user callback: <code data-editable-text-demo="callback">{lastChange === undefined ? 'No user changes' : JSON.stringify(lastChange)}</code></p>
 *       <p>Preset changes retain this control and its source. Switching to Read destroys its editor session. Dictionary changes update commands without changing content or publishing a callback.</p>
 *     </>,
 *     children: <EditableText data-editable-text-demo="preview" aria-label="Project update visible policy projection" mode={mode} value={html}
 *       fallbackText={fallback} allowedBlocks={blocks[preset]} placeholder={placeholder} strings={dictionary === 'italian' ? italianEditableTextStrings : dictionary === 'custom' ? customStrings : undefined} onChange={onChange} /> });
 * }
 *
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Edit one controlled block and compare stored HTML, visible policy projection, callback counts and live dictionaries.' };
 * ```
 */
export const EditableText = React.forwardRef<HTMLDivElement, EditableTextProps>((props, forwardedRef) => {
  validateProps(props);
  const policy = resolveEditableTextBlockPolicy(props.allowedBlocks, props.defaultBlock);
  const editing = props.mode === 'edit';
  const { targetDocument } = useFluent_unstable();
  const [root, setRoot] = React.useState<HTMLDivElement>();
  const [retry, setRetry] = React.useState(0);
  const [EditorComponent, setEditorComponent] = React.useState(() => InitialEditor);
  const [editorSession, setEditorSession] = React.useState(editing);
  const lastEditCallback = React.useRef(props.mode === 'edit' ? props.onChange : undefined);
  const renderEditor = editing || editorSession;
  React.useLayoutEffect(() => {
    if (editing) lastEditCallback.current = props.onChange;
    // Retain the session through the exit commit so the child's layout effects
    // receive the committed policy/callback before its cleanup flush. Removing
    // it synchronously here finishes Read before paint without speculative refs.
    setEditorSession(editing);
  }, [editing, props.onChange]);
  const composeRef = React.useCallback((element: HTMLDivElement | null): void => {
    // React 17 calls DOM refs with null on detach; absence in our state is undefined.
    setRoot(element ?? undefined);
    if (typeof forwardedRef === 'function') forwardedRef(element);
    else if (forwardedRef) forwardedRef.current = element;
  }, [forwardedRef]);
  // A mounted root is authoritative, including when Fluent's default context
  // captured a different global document. Without either document, mount the
  // empty root first and resolve content after its ref is attached.
  const ownerDocument = root?.ownerDocument ?? targetDocument;
  const fallbackText = props.fallbackText ?? '';
  const resolved = React.useMemo(() => ownerDocument
    ? resolveEditableTextValue(props.value, fallbackText, ownerDocument, policy) : undefined,
  [props.value, fallbackText, ownerDocument, policy.membershipKey, policy.defaultBlock]);
  // Configuration is deliberately excluded: the editor reconciles its live draft.
  const sourceKey = JSON.stringify([props.value, fallbackText]);
  const strings = React.useMemo(() => resolveEditableTextStrings(props.strings), [props.strings]);
  const publish = React.useCallback((document: EditableTextDocument): void => {
    // Read may omit its callback while a committed Edit session finishes once.
    const callback = props.onChange ?? lastEditCallback.current;
    callback?.({ html: serializeEditableTextDocument(document), text: getEditableTextPlainText(document) });
  }, [props.onChange]);
  const native = getNativeElementProps('div', props, renderEditor
    ? [...ownedProps, 'aria-label', 'aria-labelledby', 'aria-describedby'] : ownedProps);
  const read = resolved && <EditableTextReadRenderer document={resolved.document} />;
  return <div {...native} ref={composeRef}>
    {!renderEditor ? read : resolved && <>
      {resolved.kind === 'invalid' && <MessageBar intent="warning"><MessageBarBody>{strings.savedFormattingWarning}</MessageBarBody></MessageBar>}
      <React.Suspense fallback={<>{read}<Spinner size="tiny" label={strings.loadingEditor} /></>}>
        <EditableTextEditorBoundary key={retry} strings={strings} onRetry={() => {
          setEditorComponent(() => lazyEditor()); setRetry(value => value + 1);
        }}>
          <EditorComponent document={resolved.document} sourceKey={sourceKey} policy={policy} strings={strings}
            placeholder={props.placeholder ?? strings.placeholder} ariaLabel={props['aria-label']}
            ariaLabelledBy={props['aria-labelledby']} ariaDescribedBy={props['aria-describedby']} onDocumentChange={publish} />
        </EditableTextEditorBoundary>
      </React.Suspense>
    </>}
  </div>;
});
EditableText.displayName = 'EditableText';
