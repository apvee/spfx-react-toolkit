import * as React from 'react';
import { Button, Field, Input, Label, MessageBar, MessageBarBody, Select } from '@fluentui/react-components';
import { useSPFxConfigurationStore, useSPFxPageContext } from '@apvee/spfx-react-toolkit';
import type { SPFxConfigurationStoreItem, SPFxConfigurationStoreMode, SPFxConfigurationStoreScope } from '@apvee/spfx-react-toolkit';
import { ActionResult, DemoCard, InfoGrid, JsonDetails } from '../shared';
import { useDemoStyles } from '../shared/demoStyles';

type Action = 'get' | 'list' | 'save' | 'remove';
type Format = 'text' | 'json' | 'bigint' | 'date';

const ConfigurationPanel: React.FC = () => {
  const styles = useDemoStyles();
  const pageContext = useSPFxPageContext();
  const [mode, setMode] = React.useState<SPFxConfigurationStoreMode>('layered');
  const [selectedScope, setScope] = React.useState<SPFxConfigurationStoreScope>('site');
  const scope = mode === 'layered' ? selectedScope : mode;
  const store = useSPFxConfigurationStore({ mode });
  const [action, setAction] = React.useState<Action>('get');
  const [key, setKey] = React.useState('spfx-toolkit-config-demo-disposable');
  const [value, setValue] = React.useState('');
  const [format, setFormat] = React.useState<Format>('text');
  const [description, setDescription] = React.useState('Disposable configuration store demo item');
  const identity = store.get;
  const target = JSON.stringify([mode, scope, action, key, value, format, description]);
  const current = React.useRef({ identity, target });
  current.current = { identity, target };
  const mounted = React.useRef(true);
  const sequence = React.useRef(0);
  const [outcome, setOutcome] = React.useState<{
    identity: typeof identity; target: string; result?: string; error?: Error;
    item?: SPFxConfigurationStoreItem<unknown>; items?: SPFxConfigurationStoreItem<unknown>[];
  }>();
  React.useLayoutEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
  const run = async (): Promise<void> => {
    const request = ++sequence.current;
    const isCurrent = (): boolean => mounted.current && current.current.identity === identity && current.current.target === target && request === sequence.current;
    setOutcome({ identity, target });
    try {
      if (action === 'get') {
        const item = await store.get(key);
        if (isCurrent()) setOutcome({ identity, target, item, result: item ? `Found ${key}.` : `${key} not found.` });
      } else if (action === 'list') {
        const items = await store.listScope(scope);
        if (isCurrent()) setOutcome({ identity, target, items, result: `Loaded ${items.length} ${scope} physical item(s).` });
      } else if (action === 'save') {
        const input: unknown = format === 'json' ? JSON.parse(value) : format === 'bigint' ? BigInt(value) : format === 'date' ? new Date(value) : value;
        await store.save(key, input, { scope, description });
        if (isCurrent()) setOutcome({ identity, target, result: `Saved ${key} in ${scope}.` });
      } else {
        await store.remove(key, { scope });
        if (isCurrent()) setOutcome({ identity, target, result: `Removed ${key} from ${scope}. Read again to resolve inheritance.` });
      }
    } catch (failure) {
      if (isCurrent()) setOutcome({ identity, target, error: failure instanceof Error ? failure : new Error(String(failure)) });
    }
  };
  const visible = outcome?.identity === identity && outcome.target === target ? outcome : undefined;
  const busy = store.isLoading || store.isWriting;
  return <div className={styles.root}>
    <DemoCard title="Configuration Store" description="Read tenant defaults and site overrides, with an explicit destination for every change."
      iconName="TableGroup" contentPlacement="configuration"
      controls={<>
        <Field label="Mode"><Select disabled={store.isWriting} value={mode} select={{ 'data-config-store-mode': true } as React.SelectHTMLAttributes<HTMLSelectElement>}
          onChange={(_, data) => setMode(data.value as SPFxConfigurationStoreMode)}>
          <option value="tenant">Tenant only</option><option value="site">Site collection only</option><option value="layered">Tenant defaults with site overrides</option>
        </Select></Field>
        <Field label="Operation"><Select disabled={store.isWriting} value={action} select={{ 'data-demo-variant': 'ConfigurationPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>}
          onChange={(_, data) => setAction(data.value as Action)}>
          <option value="get">Read effective value</option><option value="list">Physical inventory</option><option value="save">Save</option><option value="remove">Remove</option>
        </Select></Field>
        {action !== 'get' && <Field label="Physical scope"><Select disabled={store.isWriting} value={scope} select={{ 'data-config-store-scope': true } as React.SelectHTMLAttributes<HTMLSelectElement>}
          onChange={(_, data) => setScope(data.value as SPFxConfigurationStoreScope)}>
          <option value="tenant" disabled={mode === 'site'}>Tenant</option><option value="site" disabled={mode === 'tenant'}>Site collection</option>
        </Select></Field>}
        {action !== 'list' && <Field label="Key"><Input aria-label="Key" disabled={store.isWriting} value={key} onChange={(_, data) => setKey(data.value)}/></Field>}
        {action === 'save' && <Field label="Value"><Input aria-label="Value" disabled={store.isWriting} value={value} onChange={(_, data) => setValue(data.value)}/></Field>}
      </>}
      commands={<Button appearance="primary" disabled={!store.isReady || busy || (action !== 'list' && !key)} onClick={run}>Run</Button>}
      status={<ActionResult result={visible?.result} error={visible?.error}/>}
      diagnostics={<>
        {action === 'save' && <>
          <Field label="Value format"><Select disabled={store.isWriting} value={format} select={{ 'data-config-store-format': true } as React.SelectHTMLAttributes<HTMLSelectElement>}
            onChange={(_, data) => setFormat(data.value as Format)}><option value="text">Text</option><option value="json">JSON</option><option value="bigint">Big integer</option><option value="date">Date (ISO)</option></Select></Field>
          <Field label="Description (empty clears destination metadata)"><Input aria-label="Description" disabled={store.isWriting} value={description} onChange={(_, data) => setDescription(data.value)}/></Field>
        </>}
        <MessageBar intent="info"><MessageBarBody>
          Read checks the site first in layered mode; only confirmed absence uses tenant defaults. Objects are replaced whole.
          Physical inventory lists one store without merging. Only Save may initialize its selected store.
          Use unique spfx-toolkit-config-demo- keys and remove them from every scope after testing.
        </MessageBarBody></MessageBar>
        {visible?.item && <JsonDetails label="Resolved configuration record" value={visible.item}/>}
        {visible?.items && <JsonDetails label="Physical configuration records" value={visible.items}/>}
      </>}>
      {busy && <Label>Loading...</Label>}
      <InfoGrid rows={[
        { label: 'Collection root URL', value: pageContext.site.absoluteUrl },
        { label: 'Current web URL', value: pageContext.web.absoluteUrl },
        { label: 'Source', value: visible?.item?.source },
        { label: 'Selected key', value: visible?.item?.key },
        { label: 'Selected value', value: visible?.item ? String(visible.item.value) : undefined },
        { label: 'Value type', value: visible?.item ? typeof visible.item.value : undefined },
      ]}/>
      {visible?.items && <InfoGrid rows={visible.items.slice(0, 10).map(item => ({ label: item.key, value: String(item.value) }))}/>}
    </DemoCard>
  </div>;
};
export default ConfigurationPanel;
