import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem, Button, Text } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import type { ExecutableExampleEntry } from '../generated/jsdoc-examples/catalogue';
import { DemoAccordionPanel } from './DemoAccordionPanel';
import { useExampleExplorerStyles } from './exampleExplorerStyles';

type SourceState = { readonly status: 'idle' | 'pending' | 'error' } | { readonly status: 'ready'; readonly code: string };
type CopyState = 'idle' | 'pending' | 'copied' | 'error';

/** Immutable authored code, owned by the selected explorer mount rather than its demo Reset. */
export function ExecutableExampleSource({ entry }: { readonly entry: ExecutableExampleEntry }): React.ReactElement {
  const styles = useExampleExplorerStyles();
  const { targetDocument } = useFluent_unstable();
  const [open, setOpen] = React.useState(false);
  const [source, setSource] = React.useState<SourceState>({ status: 'idle' });
  const [copy, setCopy] = React.useState<CopyState>('idle');
  const mounted = React.useRef(true);
  const copying = React.useRef(false);
  React.useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; };
  }, []);

  const load = (): void => {
    setSource({ status: 'pending' });
    // Catch synchronous loader failures as well as rejected chunk requests.
    Promise.resolve().then(() => entry.loadSource()).then(module => {
      if (mounted.current) setSource({ status: 'ready', code: module.code });
    }, () => {
      if (mounted.current) setSource({ status: 'error' });
    });
  };
  const copyCode = (): void => {
    if (source.status !== 'ready' || copying.current) return;
    const code = source.code;
    copying.current = true;
    setCopy('pending');
    const write = async (): Promise<void> => {
      const clipboard = targetDocument?.defaultView?.navigator.clipboard;
      if (!clipboard) throw new Error('Clipboard unavailable');
      await clipboard.writeText(code);
    };
    write().then(() => {
      if (mounted.current) { copying.current = false; setCopy('copied'); }
    }, () => {
      if (mounted.current) { copying.current = false; setCopy('error'); }
    });
  };
  const message = source.status === 'pending' ? 'Loading example code...'
    : source.status === 'error' ? 'Unable to load example code.'
    : copy === 'pending' ? 'Copying...' : copy === 'copied' ? 'Copied'
    : copy === 'error' ? 'Unable to copy code. Select and copy it manually.' : '';

  return <Accordion collapsible openItems={open ? ['source'] : []} className={styles.source}
    onToggle={(_event, data) => {
      const nextOpen = data.openItems.includes('source');
      setOpen(nextOpen);
      if (nextOpen && source.status === 'idle') load();
    }}>
    <AccordionItem value="source">
      <AccordionHeader>Example code</AccordionHeader>
      <DemoAccordionPanel className={styles.sourcePanel}>
        <Text className={styles.sourceProvenance}>{entry.source.file}:{entry.source.line}</Text>
        <div className={styles.toolbar}>
          <Button disabled={source.status !== 'ready' || copy === 'pending'} onClick={copyCode}>{copy === 'pending' ? 'Copying...' : 'Copy code'}</Button>
          {source.status === 'error' && <Button onClick={load}>Retry</Button>}
        </div>
        <Text role="status" aria-live="polite">{message}</Text>
        {source.status === 'ready' && <pre data-example-source-code className={styles.sourceCode} role="region" aria-label="Example code" tabIndex={0} dir="ltr"><code>{source.code}</code></pre>}
      </DemoAccordionPanel>
    </AccordionItem>
  </Accordion>;
}
