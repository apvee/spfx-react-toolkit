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
 * @example
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
