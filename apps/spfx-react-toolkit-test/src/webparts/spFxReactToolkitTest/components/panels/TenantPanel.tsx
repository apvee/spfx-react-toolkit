import * as React from 'react';
import { Select, Button, Field, Input, MessageBar, MessageBarBody, Label } from '@fluentui/react-components';
import { useSPFxTenantKeyValueStore, useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';
import type { SPFxTenantKeyValueStoreItem } from '@apvee/spfx-react-toolkit';
import { ActionResult, DemoCard, InfoGrid, JsonDetails, StatusBadge } from '../shared';
import { useDemoStyles } from '../shared/demoStyles';
type TenantAction = 'list' | 'get' | 'save' | 'remove';
const TenantPropertyScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const version = useSPFxTenantProperty<string>('spfx-toolkit-test-version', false);
    const counter = useSPFxTenantProperty<number>('spfx-toolkit-test-counter', false);
    const [executed, setExecuted] = React.useState<{ readonly version: typeof version.load; readonly counter: typeof counter.load; readonly request: number }>();
    const generation = React.useRef(0);
    const identity = React.useRef({ version: version.load, counter: counter.load });
    if (identity.current.version !== version.load || identity.current.counter !== counter.load) {
        identity.current = { version: version.load, counter: counter.load };
        generation.current += 1;
    }
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const execute = React.useCallback(async () => {
        const snapshot = { version: version.load, counter: counter.load, request: ++generation.current };
        setExecuted(undefined);
        await Promise.all([snapshot.version(), snapshot.counter()]);
        if (mounted.current && generation.current === snapshot.request && identity.current.version === snapshot.version && identity.current.counter === snapshot.counter)
            setExecuted(snapshot);
    }, [version, counter]);
    const current = executed?.request === generation.current && executed.version === version.load && executed.counter === counter.load;
    const loading = version.isLoading || counter.isLoading;
    const error = version.error ?? counter.error;
    return <DemoCard title="Tenant Properties" description="Read the tenant version and counter." iconName="ServerEnviroment" controls={selector()} error={error} contentPlacement="configuration" commands={<Button appearance="primary" disabled={loading} onClick={execute}>Run</Button>} diagnostics={<>
        <Label>Tenant properties are read-only through REST.</Label><InfoGrid rows={[{ label: 'Version Description', value: version.description }, { label: 'Counter Description', value: counter.description }]}/></>}>
    {loading ? <Label>Loading...</Label> : !current || error ? <Label>Run to read the properties.</Label> : <InfoGrid rows={[{ label: 'spfx-toolkit-test-version', value: version.data ?? '(not set)' }, { label: 'spfx-toolkit-test-counter', value: counter.data === undefined ? '(not set)' : String(counter.data) }]}/>}
  </DemoCard>;
};
const TenantStoreScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const store = useSPFxTenantKeyValueStore();
    const [action, setAction] = React.useState<TenantAction>('get');
    const [key, setKey] = React.useState('spfx-toolkit-tenant-demo-disposable');
    const [value, setValue] = React.useState('');
    const [description, setDescription] = React.useState('Disposable SPFx React Toolkit site demo item');
    // get changes when the hook's collection or client changes, including in-place PageContext changes.
    const identity = store.get;
    const target = `${action}:${key}:${value}:${description}`;
    const currentTarget = React.useRef(target);
    currentTarget.current = target;
    const currentIdentity = React.useRef(identity);
    currentIdentity.current = identity;
    const mounted = React.useRef(true);
    const request = React.useRef(0);
    const [outcome, setOutcome] = React.useState<{
        identity: typeof identity;
        target: string;
        action: TenantAction;
        result?: string;
        error?: Error;
        selectedItem?: SPFxTenantKeyValueStoreItem<unknown>;
        items?: SPFxTenantKeyValueStoreItem<unknown>[];
    }>();
    React.useLayoutEffect(() => {
        mounted.current = true;
        return () => { mounted.current = false; };
    }, []);
    const runAction = React.useCallback(async (action: TenantAction): Promise<void> => {
        const operation = ++request.current;
        const isCurrent = (): boolean => mounted.current && currentIdentity.current === identity && request.current === operation && currentTarget.current === target;
        setOutcome({ identity, target, action });
        try {
            if (action === 'list') {
                const items = await store.list();
                if (isCurrent())
                    setOutcome({ identity, target, action, items, result: `Loaded ${items.length} tenant key-value item(s).` });
            }
            else if (action === 'get') {
                const selectedItem = await store.get<string>(key);
                if (isCurrent())
                    setOutcome({ identity, target, action, selectedItem, result: selectedItem ? `Found ${key}.` : `${key} not found.` });
            }
            else if (action === 'save') {
                await store.save<string>(key, value, description || undefined);
                if (isCurrent())
                    setOutcome({ identity, target, action, result: `Saved ${key}.` });
            }
            else {
                await store.remove(key);
                if (isCurrent())
                    setOutcome({ identity, target, action, result: `Removed ${key}.` });
            }
        }
        catch (failure) {
            if (isCurrent())
                setOutcome({ identity, target, action, error: failure instanceof Error ? failure : new Error(String(failure)) });
        }
    }, [identity, store, key, value, description, target]);
    const currentOutcome = outcome?.identity === identity && outcome.target === target ? outcome : undefined;
    const isRead = currentOutcome?.action === 'get' || currentOutcome?.action === 'list';
    // Reads resolve fallback values on failure. Only current rendered error state can distinguish failure from absence.
    const error = currentOutcome ? currentOutcome.error ?? (isRead ? store.error : store.writeError) : undefined;
    const result = error ? undefined : currentOutcome?.result;
    const busy = store.isLoading || store.isWriting;
    return (<div className={demoStyles.root}>
    <DemoCard title="Tenant Key-Value Store" description="Read and write disposable values shared by this tenant." status={<ActionResult result={result} error={error}/>} iconName="TableGroup" controls={<>{selector(store.isWriting)}<Field label="Operation"><Select disabled={store.isWriting} select={{ 'data-demo-variant': 'TenantStoreScenario' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={action} onChange={(_, data) => setAction(data.value as TenantAction)}><option value="get">Read</option><option value="list">List items</option><option value="save">Save</option><option value="remove">Remove</option></Select></Field>
      {action !== 'list' && <Field label={"Key"}>
        <Input disabled={store.isWriting} value={key} onChange={(_, data) => setKey(data.value ?? '')}/>
      </Field>}
      {action === 'save' && <Field label={"Value"}>
        <Input disabled={store.isWriting} value={value} onChange={(_, data) => setValue(data.value ?? '')}/>
      </Field>}
      {action === 'save' && <Field label={"Description"}>
        <Input disabled={store.isWriting} value={description} onChange={(_, data) => setDescription(data.value ?? '')}/>
      </Field>}

    </>} contentPlacement="configuration" commands={<>
      <Button appearance="primary" disabled={!store.isReady || busy || (action !== 'list' && !key)} onClick={() => runAction(action)}>Run</Button>
    </>} diagnostics={<>
      <div className={demoStyles.actions}>
        <StatusBadge label="Ready" available={store.isReady}/>
        <StatusBadge label="Can Write" available={store.canWrite}/>
        <StatusBadge label="Loading" available={!store.isLoading}/>
        <StatusBadge label="Writing" available={!store.isWriting}/>
      </div>
      <MessageBar intent="info">
        <MessageBarBody>
          All subsites share SiteKeyValueStore in the tenant. Actions run only when clicked.
          Use disposable keys beginning with spfx-toolkit-tenant-demo- and remove them after testing.
          Saving may create the store; SharePoint enforces permissions. Can Write is an indicator.
        </MessageBarBody>
      </MessageBar>
      {!error && currentOutcome?.selectedItem && <JsonDetails label="Selected tenant key-value item" value={currentOutcome.selectedItem}/>}
      {!error && currentOutcome?.items && currentOutcome.items.length > 0 && <JsonDetails label="Site key-value items" value={currentOutcome.items}/>}
    </>}>

      {busy && <Label>Loading...</Label>}
      {!busy && !error && currentOutcome?.items && <InfoGrid rows={currentOutcome.items.slice(0, 10).map(item => ({ label: item.key, value: String(item.value) }))}/>}
      <InfoGrid rows={[
            { label: 'Selected key', value: !error ? currentOutcome?.selectedItem?.key : undefined },
            { label: 'Selected value', value: !error && currentOutcome?.selectedItem ? String(currentOutcome.selectedItem.value) : undefined },
        ]}/>


    </DemoCard>
  </div>);
};
const TenantPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('properties');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Service"><Select disabled={busy} select={{ 'data-demo-variant': 'TenantPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}><option value="properties">Tenant properties</option><option value="store">Tenant storage</option></Select></Field>;
    return scenario === 'properties' ? <TenantPropertyScenario selector={selector}/> : <TenantStoreScenario selector={selector}/>;
};
export default TenantPanel;
