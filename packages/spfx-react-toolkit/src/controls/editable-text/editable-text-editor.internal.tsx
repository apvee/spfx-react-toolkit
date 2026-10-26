import * as React from 'react';
import { Editor } from '@tiptap/core';
import { EditorContent } from '@tiptap/react';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';
import { MessageBar, MessageBarBody } from '@fluentui/react-components';
import { createEditableTextExtensions } from './editable-text-extensions.internal';
import { editorContentToEditableTextDocument, editableTextDocumentToEditorContent } from './editable-text-editor-adapter.internal';
import { resolveEditableTextValue, serializeEditableTextDocument } from './editable-text-codec.internal';
import { normalizeEditableTextDocument } from './editable-text-policy.internal';
import type { EditableTextBlockPolicy } from './editable-text-policy.internal';
import { useEditableTextStyles } from './editable-text-styles.internal';
import { EditableTextToolbar } from './editable-text-toolbar.internal';
import { resolveEditableTextStrings } from './localization/editable-text-locales.internal';
import type { EditableTextStrings } from './localization/editable-text-strings.internal';
import type { EditableTextDocument } from './editable-text-document.types.internal';

const useEditorStyles = makeStyles({ root: {
  '& .tiptap': { minInlineSize: 0 },
  '& .is-editor-empty': { position: 'relative' },
  '& .is-editor-empty::before': { content: 'attr(data-placeholder)', color: tokens.colorNeutralForeground3,
    position: 'absolute', insetInlineStart: 0, pointerEvents: 'none' }
} });

/** Private lazy boundary. Engine types remain confined to this implementation. */
export interface EditableTextEditorProps {
  document: EditableTextDocument;
  /** Identity of the controlled value/fallback pair, independent of configuration. */
  sourceKey: string;
  policy: EditableTextBlockPolicy;
  strings?: EditableTextStrings;
  placeholder: string;
  ariaLabel?: string;
  ariaLabelledBy?: string;
  ariaDescribedBy?: string;
  onDocumentChange: (document: EditableTextDocument) => void;
}
function documentKey(document: EditableTextDocument): string {
  try { return serializeEditableTextDocument(document); }
  catch (error) { if (error instanceof RangeError) return JSON.stringify(document); throw error; }
}
/** Preserve a deliberate allowed empty block and the stored marks prepared for typing. */
function normalizeLive(document: EditableTextDocument, policy: EditableTextBlockPolicy): EditableTextDocument {
  const block = document.type === 'paragraph' ? 'p' : `h${document.level}` as const;
  return policy.blocks.includes(block) ? document : normalizeEditableTextDocument(document, policy);
}
function attributes(props: EditableTextEditorProps): Record<string, string> {
  return { ...(props.ariaLabel ? { 'aria-label': props.ariaLabel } : {}),
    ...(props.ariaLabelledBy ? { 'aria-labelledby': props.ariaLabelledBy } : {}),
    ...(props.ariaDescribedBy ? { 'aria-describedby': props.ariaDescribedBy } : {}),
    role: 'textbox', 'aria-multiline': 'true', spellcheck: 'true' };
}

/** Stable Editor/View/schema; source replacement and policy reconciliation have distinct history semantics. */
export function EditableTextEditor(props: EditableTextEditorProps): React.ReactElement {
  const strings = props.strings ?? resolveEditableTextStrings();
  const styles = useEditableTextStyles();
  const editorStyles = useEditorStyles();
  const latest = React.useRef(props);
  const wrapper = React.useRef<HTMLDivElement>(null); // React 17 DOM refs require the legacy null sentinel.
  const reconcile = React.useRef<(sourceChanged: boolean) => void>();
  const instance = React.useRef<Editor>();
  const [editor, setEditor] = React.useState<Editor>();
  const [error, setError] = React.useState(false);
  const syncing = React.useRef({ active: false, dirty: false, composing: false, revision: 0, compositionRevision: 0,
    published: documentKey(props.document), membership: props.policy.membershipKey, defaultBlock: props.policy.defaultBlock,
    pending: undefined as EditableTextDocument | undefined, policyPending: false });
  React.useLayoutEffect(() => { latest.current = props; });

  React.useLayoutEffect(() => {
    const state = syncing.current;
    state.active = true;
    const replace = (document: EditableTextDocument, source: boolean): void => {
      const current = instance.current;
      if (!current || current.isDestroyed) return;
      const node = current.schema.nodeFromJSON(editableTextDocumentToEditorContent(document));
      const max = Math.max(1, node.content.size - 1);
      const clamp = (position: number): number => Math.max(1, Math.min(position, max));
      const selection = TextSelection.create(node, clamp(current.state.selection.anchor), clamp(current.state.selection.head));
      // A changed membership establishes a fresh history base even when the block
      // remains allowed. Keep typing marks; external source changes remain authoritative.
      current.view.updateState(EditorState.create({ schema: current.schema, doc: node, selection,
        storedMarks: source ? undefined : current.state.storedMarks, plugins: current.state.plugins }));
      if (source) { state.published = documentKey(document); state.dirty = false; setError(false); }
      current.emit('selectionUpdate', { editor: current, transaction: current.state.tr });
    };
    const enforce = (): EditableTextDocument | undefined => {
      const current = instance.current;
      if (!current || current.isDestroyed) return;
      const live = editorContentToEditableTextDocument(current.getJSON());
      const policy = latest.current.policy;
      const defaultOnly = !state.policyPending && state.defaultBlock !== policy.defaultBlock;
      const normalized = defaultOnly ? normalizeEditableTextDocument(live, policy) : normalizeLive(live, policy);
      if (state.policyPending) replace(normalized, false);
      else if (normalized !== live) {
        // A default-only update changes an empty block without discarding undo.
        // This private configuration transaction is never a local publication.
        const block = current.schema.nodeFromJSON(editableTextDocumentToEditorContent(normalized)).firstChild;
        if (block) {
          const tr = current.state.tr.setNodeMarkup(0, block.type, block.attrs)
            .setMeta('addToHistory', false).setMeta('preventUpdate', true);
          if (current.state.storedMarks) tr.setStoredMarks(current.state.storedMarks);
          current.view.dispatch(tr);
        }
      }
      state.membership = policy.membershipKey; state.defaultBlock = policy.defaultBlock; state.policyPending = false;
      return normalized;
    };
    const publish = (disposing = false): void => {
      const current = instance.current;
      if (!state.active || !current || current.isDestroyed) return;
      // Blur/unmount may precede compositionend; the final flush always uses the
      // latest policy and live text, including a draft that exceeds the byte budget.
      const document = enforce();
      if (!state.dirty || !document) return;
      try {
        const canonical = serializeEditableTextDocument(document);
        if (canonical !== state.published) {
          state.published = canonical;
          latest.current.onDocumentChange(document);
        }
        state.dirty = false;
        if (!disposing) setError(false);
      } catch (caught) {
        if (!(caught instanceof RangeError)) throw caught;
        if (!disposing) setError(true);
      }
    };
    const current = new Editor({
      element: wrapper.current?.ownerDocument.createElement('div'),
      extensions: createEditableTextExtensions(() => latest.current.placeholder, () => latest.current.policy,
        () => state.composing || instance.current?.view.composing === true),
      content: editableTextDocumentToEditorContent(normalizeEditableTextDocument(latest.current.document, latest.current.policy)), injectCSS: false,
      editorProps: { attributes: attributes(latest.current), handleDOMEvents: {
        compositionstart: () => { state.composing = true; state.compositionRevision = state.revision; return false; },
        compositionend: () => {
          // ProseMirror flushes its DOM observer after this handler. Arbitrate
          // after that flush; no debounce timer or editor recreation is involved.
          queueMicrotask(() => queueMicrotask(() => {
            if (!state.active) return;
            state.composing = false;
            const pending = state.pending; state.pending = undefined;
            if (pending && state.revision === state.compositionRevision) {
              replace(normalizeEditableTextDocument(pending, latest.current.policy), true);
              state.membership = latest.current.policy.membershipKey; state.defaultBlock = latest.current.policy.defaultBlock; state.policyPending = false;
            } else publish();
          }));
          return false;
        },
        blur: () => { publish(); return false; }
      }, handlePaste: (_view, event) => {
        const clipboard = event.clipboardData;
        if (!clipboard) return false;
        const html = clipboard.getData('text/html'); const text = clipboard.getData('text/plain');
        const policy = latest.current.policy;
        const imported: EditableTextDocument = html ? resolveEditableTextValue(html, '', current.view.dom.ownerDocument, policy).document : {
          type: 'paragraph', content: text.split(/\r?\n/).flatMap((line, index) => [
            ...(index ? [{ type: 'hardBreak' as const }] : []), ...(line ? [{ type: 'text' as const, text: line }] : []) ]) };
        const document = normalizeEditableTextDocument(imported, policy);
        const block = editableTextDocumentToEditorContent(document).content?.[0];
        if (!block) return true;
        current.chain().setNode(document.type, block.attrs).insertContent(block.content ?? []).run();
        return true;
      } },
      onUpdate: () => {
        if (!state.active) return;
        state.dirty = true; state.revision++;
        if (!state.composing && !current.view.composing) publish();
      }, onBlur: () => publish()
    });
    instance.current = current;
    reconcile.current = sourceChanged => {
      const policy = latest.current.policy;
      if (state.membership !== policy.membershipKey) state.policyPending = true;
      if (sourceChanged) {
        const source = normalizeEditableTextDocument(latest.current.document, policy);
        if (documentKey(source) === documentKey(editorContentToEditableTextDocument(current.getJSON()))) state.pending = undefined;
        else if (state.composing || current.view.composing) state.pending = source;
        else replace(source, true);
      }
      if (!state.composing && !current.view.composing) enforce();
    };
    setEditor(current);
    return () => { publish(true); state.active = false; reconcile.current = undefined; instance.current = undefined; current.destroy(); };
  }, []);

  React.useLayoutEffect(() => {
    const current = instance.current;
    if (!current || current.isDestroyed) return;
    current.setOptions({ editorProps: { ...current.options.editorProps, attributes: attributes(props) } });
    current.view.dispatch(current.state.tr.setMeta('addToHistory', false));
  }, [props.ariaLabel, props.ariaLabelledBy, props.ariaDescribedBy, props.placeholder]);
  const previousSource = React.useRef(props.sourceKey);
  React.useLayoutEffect(() => {
    const sourceChanged = previousSource.current !== props.sourceKey;
    previousSource.current = props.sourceKey;
    reconcile.current?.(sourceChanged);
  }, [props.sourceKey, props.policy.membershipKey, props.policy.defaultBlock]);
  return <div className={[styles.root, editorStyles.root].join(' ')} ref={wrapper}>
    {error && <MessageBar intent="error"><MessageBarBody>{strings.tooLargeError}</MessageBarBody></MessageBar>}
    {editor && <EditableTextToolbar editor={editor} policy={props.policy} ownerDocument={editor.view.dom.ownerDocument} strings={strings} />}
    {/* Tiptap React requires its legacy null sentinel. */}
    <EditorContent editor={editor ?? null} />
  </div>;
}
