import * as React from 'react';
import { Button, MessageBar, MessageBarBody, Spinner } from '@fluentui/react-components';
import { getTitlePlainText, resolveWebPartTitle, serializeTitleDocument } from './web-part-title-codec.internal';
import { WebPartTitleReadRenderer } from './web-part-title-renderer.internal';
import { resolveWebPartShellStrings } from '../localization/web-part-shell-locales.internal';
import type { WebPartShellStrings } from '../localization/web-part-shell-strings.internal';
import type { TitleDocument } from './web-part-title.types.internal';

/** Controlled private title boundary; only the Shell owns persistence. */
export interface WebPartTitleProps {
  strings?: WebPartShellStrings;
  title: string;
  titleHTML?: string;
  showTitle: boolean;
  editing: boolean;
  onChange: (update: { title: string; titleHTML: string }) => void;
  ownerDocument: Document;
}

const loadEditor = (): Promise<typeof import('./web-part-title-editor.internal')> => import(/* webpackChunkName: 'web-part-title-editor' */ './web-part-title-editor.internal');
function lazyEditor(): React.LazyExoticComponent<typeof import('./web-part-title-editor.internal')['WebPartTitleEditor']> {
  return React.lazy(() => loadEditor().then(module => ({ default: module.WebPartTitleEditor })));
}
// Stable lazy identity; invoking the loader occurs only when this component is rendered in Edit.
const InitialEditor = lazyEditor();

class TitleEditorBoundary extends React.Component<{ children: React.ReactNode; onRetry: () => void; strings: WebPartShellStrings }, { failed: boolean }> {
  public state = { failed: false };
  public static getDerivedStateFromError(): { failed: boolean } { return { failed: true }; }
  public render(): React.ReactNode {
    return this.state.failed ? <MessageBar intent="error"><MessageBarBody>{this.props.strings.titleEditorLoadError}</MessageBarBody>
      <Button onClick={this.props.onRetry}>{this.props.strings.retry}</Button></MessageBar> : this.props.children;
  }
}

/** Read requires only codec/renderer; recover engine failures locally without wrapping app children. */
export function WebPartTitle(props: WebPartTitleProps): React.ReactElement {
  const strings = props.strings ?? resolveWebPartShellStrings(undefined);
  const [retry, setRetry] = React.useState(0);
  const [EditorComponent, setEditorComponent] = React.useState(() => InitialEditor);
  const resolved = React.useMemo(() => resolveWebPartTitle(props.title, props.titleHTML, props.ownerDocument),
    [props.title, props.titleHTML, props.ownerDocument]);
  const publish = React.useCallback((document: TitleDocument): void => {
    const titleHTML = serializeTitleDocument(document);
    props.onChange({ title: getTitlePlainText(document), titleHTML });
  }, [props.onChange]);
  if (!props.showTitle) return <></>;
  const read = <WebPartTitleReadRenderer document={resolved.document} />;
  if (!props.editing) return read;
  return <>
    {resolved.kind === 'invalid' && <MessageBar intent="warning"><MessageBarBody>{strings.savedTitleFormattingWarning}</MessageBarBody></MessageBar>}
    <React.Suspense fallback={<>{read}<Spinner size="tiny" label={strings.loadingTitleEditor} /></>}>
      <TitleEditorBoundary strings={strings} key={retry} onRetry={() => { setEditorComponent(() => lazyEditor()); setRetry(value => value + 1); }}>
        <EditorComponent document={resolved.document} onDocumentChange={publish} strings={strings} ariaLabel={strings.titleEditorLabel} placeholder={strings.titlePlaceholder} />
      </TitleEditorBoundary>
    </React.Suspense>
  </>;
}
