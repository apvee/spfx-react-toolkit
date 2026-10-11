import * as React from 'react';
import { Editor } from '@tiptap/core';
import { EditorContent } from '@tiptap/react';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';
import { MessageBar, MessageBarBody } from '@fluentui/react-components';
import { createTitleExtensions } from './web-part-title-extensions.internal';
import { editorContentToTitleDocument, titleDocumentToEditorContent } from './web-part-title-editor-adapter.internal';
import { parseTitleHTML, serializeTitleDocument } from './web-part-title-codec.internal';
import { useWebPartTitleStyles } from './web-part-title-styles.internal';
import { WebPartTitleToolbar } from './web-part-title-toolbar.internal';
import type { TitleDocument } from './web-part-title.types.internal';

const useEditorStyles = makeStyles({ root: {
  '& .tiptap': { minInlineSize: 0 },
  '& .is-editor-empty': { position: 'relative' },
  '& .is-editor-empty::before': { content: 'attr(data-placeholder)', color: tokens.colorNeutralForeground3,
    position: 'absolute', insetInlineStart: 0, pointerEvents: 'none' }
} });

/** Private async boundary; engine types are never exposed from the public control. */
export interface WebPartTitleEditorProps {
  document: TitleDocument;
  onDocumentChange: (document: TitleDocument) => void;
  ariaLabel: string;
  placeholder: string;
}

function documentKey(document: TitleDocument): string {
  try { return serializeTitleDocument(document); }
  catch (error) { if (error instanceof RangeError) return JSON.stringify(document); throw error; }
}

/** One stable view per Edit session with live atomic publication and authoritative replacement. */
export function WebPartTitleEditor(props: WebPartTitleEditorProps): React.ReactElement {
  const styles = useWebPartTitleStyles();
  const editorStyles = useEditorStyles();
  const latest = React.useRef(props);
  const wrapper = React.useRef<HTMLDivElement>(null); // React 17 DOM refs use the legacy null sentinel.
  const externalReplacement = React.useRef<((document: TitleDocument) => void)>();
  const instance = React.useRef<Editor>();
  const [editor, setEditor] = React.useState<Editor>();
  const [error, setError] = React.useState<string>();
  const syncing = React.useRef({ active: false, dirty: false, composing: false, revision: 0, compositionRevision: 0,
    published: documentKey(props.document), pending: undefined as TitleDocument | undefined });
  React.useLayoutEffect(() => { latest.current = props; });

  React.useLayoutEffect(() => {
    const state = syncing.current;
    state.active = true;
    const publish = (disposing = false): void => {
      const current = instance.current;
      if (!state.active || !state.dirty || !current || current.isDestroyed) return;
      try {
        const document = editorContentToTitleDocument(current.getJSON());
        const canonical = serializeTitleDocument(document);
        if (canonical !== state.published) {
          state.published = canonical;
          latest.current.onDocumentChange(document);
        }
        state.dirty = false;
        if (!disposing) setError(undefined);
      } catch (caught) {
        if (!(caught instanceof RangeError)) throw caught;
        // Keep both properties untouched and preserve the draft until it fits the protocol.
        if (!disposing) setError('Title is too large to save. Shorten it to continue.');
      }
    };
    const replace = (document: TitleDocument): void => {
      const current = instance.current;
      if (!current || current.isDestroyed) return;
      const node = current.schema.nodeFromJSON(titleDocumentToEditorContent(document));
      const max = Math.max(1, node.content.size - 1);
      const clampInline = (position: number): number => Math.max(1, Math.min(position, max));
      // Deliberately convert AllSelection/NodeSelection to an inline text range: typing
      // must replace title text while retaining the authoritative single-block format.
      const selection = TextSelection.create(node, clampInline(current.state.selection.anchor), clampInline(current.state.selection.head));
      // Fresh plugin state deliberately resets history while preserving the existing EditorView.
      current.view.updateState(EditorState.create({ schema: current.schema, doc: node, selection, plugins: current.state.plugins }));
      state.published = documentKey(document); state.dirty = false;
      setError(undefined);
      current.emit('selectionUpdate', { editor: current, transaction: current.state.tr });
    };
    const current = new Editor({
      element: wrapper.current?.ownerDocument.createElement('div'),
      extensions: createTitleExtensions(latest.current.placeholder),
      content: titleDocumentToEditorContent(latest.current.document), injectCSS: false,
      editorProps: { attributes: { 'aria-label': latest.current.ariaLabel, role: 'textbox', 'aria-multiline': 'true', spellcheck: 'true' },
        handleDOMEvents: {
          compositionstart: () => { state.composing = true; state.compositionRevision = state.revision; return false; },
          compositionend: () => {
            // ProseMirror schedules its final DOM-observer flush in a microtask after
            // this handler returns. Arbitrate after that flush, without a timer/debounce.
            queueMicrotask(() => queueMicrotask(() => {
              if (!state.active) return;
              state.composing = false;
              const pending = state.pending; state.pending = undefined;
              if (pending && state.revision === state.compositionRevision) replace(pending);
              else publish();
            }));
            return false;
          },
          blur: () => { publish(); return false; }
        },
        handlePaste: (_view, event) => {
          const clipboard = event.clipboardData;
          if (!clipboard) return false;
          const html = clipboard.getData('text/html');
          const text = clipboard.getData('text/plain');
          const document: TitleDocument = html ? parseTitleHTML(html, current.view.dom.ownerDocument).document : {
            type: 'paragraph', content: text.split(/\r?\n/).flatMap((line, index) => [
              ...(index ? [{ type: 'hardBreak' as const }] : []), ...(line ? [{ type: 'text' as const, text: line }] : []) ]) };
          const block = titleDocumentToEditorContent(document).content?.[0];
          if (!block) return true;
          current.chain().setNode(document.type, block.attrs).insertContent(block.content ?? []).run();
          return true;
        }
      },
      onUpdate: () => {
        if (!state.active) return;
        state.dirty = true; state.revision++;
        if (!state.composing) publish();
      },
      onBlur: () => publish()
    });
    instance.current = current;
    // Keep synchronization functions private to this instance, without an editor factory/registry.
    externalReplacement.current = document => {
      if (state.composing || current.view.composing) state.pending = document;
      else replace(document);
    };
    setEditor(current);
    return () => { publish(true); state.active = false; externalReplacement.current = undefined; instance.current = undefined; current.destroy(); };
  }, []);

  React.useEffect(() => {
    const current = instance.current;
    if (!current || current.isDestroyed) return;
    if (documentKey(props.document) === documentKey(editorContentToTitleDocument(current.getJSON()))) {
      // A return to the current value also supersedes an earlier queued external title,
      // without replacing the view state or resetting its selection/history.
      syncing.current.pending = undefined;
      return;
    }
    externalReplacement.current?.(props.document);
  }, [props.document]);
  return <div className={[styles.root, editorStyles.root].join(' ')} ref={wrapper}>
    {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
    {editor && <WebPartTitleToolbar editor={editor} ownerDocument={editor.view.dom.ownerDocument} />}
    {/* Tiptap React requires the legacy null sentinel; internal absence remains undefined. */}
    <EditorContent editor={editor ?? null} />
  </div>;
}
