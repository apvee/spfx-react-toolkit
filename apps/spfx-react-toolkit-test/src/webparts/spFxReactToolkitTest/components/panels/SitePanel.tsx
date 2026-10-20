import { useDemoStyles } from '../shared/demoStyles';
import { Select, Button, Field, Input, MessageBar, MessageBarBody, Label } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPageContext, useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit';
import type { SPFxSiteKeyValueStoreItem } from '@apvee/spfx-react-toolkit';
import { ActionResult, DemoCard, InfoGrid, JsonDetails, StatusBadge } from '../shared';
type SiteAction = 'list' | 'get' | 'save' | 'remove';
const SitePanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const pageContext = useSPFxPageContext();
    const store = useSPFxSiteKeyValueStore();
    const [action, setAction] = React.useState<SiteAction>('get');
    const [key, setKey] = React.useState('spfx-toolkit-site-demo-disposable');
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
        action: SiteAction;
        result?: string;
        error?: Error;
        selectedItem?: SPFxSiteKeyValueStoreItem<unknown>;
        items?: SPFxSiteKeyValueStoreItem<unknown>[];
    }>();
    React.useLayoutEffect(() => {
        mounted.current = true;
        return () => { mounted.current = false; };
    }, []);
    const runAction = React.useCallback(async (action: SiteAction): Promise<void> => {
        const operation = ++request.current;
        const isCurrent = (): boolean => mounted.current && currentIdentity.current === identity && request.current === operation && currentTarget.current === target;
        setOutcome({ identity, target, action });
        try {
            if (action === 'list') {
                const items = await store.list();
                if (isCurrent())
                    setOutcome({ identity, target, action, items, result: `Loaded ${items.length} site key-value item(s).` });
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
    <DemoCard title="Site Key-Value Store" description="Read and write disposable values shared by this site collection." status={<ActionResult result={result} error={error}/>} iconName="TableGroup" controls={<><Field className={demoStyles.controlRow} label="Operation"><Select disabled={store.isWriting} select={{ 'data-demo-variant': 'SitePanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={action} onChange={(_, data) => setAction(data.value as SiteAction)}><option value="get">Read</option><option value="list">List</option><option value="save">Save</option><option value="remove">Remove</option></Select></Field>
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
          All subsites share SiteKeyValueStore in the collection root. Actions run only when clicked.
          Use disposable keys beginning with spfx-toolkit-site-demo- and remove them after testing.
          Saving may create the store; SharePoint enforces permissions. Can Write is an indicator.
        </MessageBarBody>
      </MessageBar>
      {!error && currentOutcome?.selectedItem && <JsonDetails label="Selected site key-value item" value={currentOutcome.selectedItem}/>}
      {!error && currentOutcome?.items && currentOutcome.items.length > 0 && <JsonDetails label="Site key-value items" value={currentOutcome.items}/>}
    </>}>

      {busy && <Label>Loading...</Label>}
      {!busy && !error && currentOutcome?.items && <InfoGrid rows={currentOutcome.items.slice(0, 10).map(item => ({ label: item.key, value: String(item.value) }))}/>}
      <InfoGrid rows={[
            { label: 'Collection root URL', value: pageContext.site.absoluteUrl, icon: 'Info' },
            { label: 'Current web URL', value: pageContext.web.absoluteUrl, icon: 'Info' },
            { label: 'Selected key', value: !error ? currentOutcome?.selectedItem?.key : undefined },
            { label: 'Selected value', value: !error && currentOutcome?.selectedItem ? String(currentOutcome.selectedItem.value) : undefined },
        ]}/>


    </DemoCard>
  </div>);
};
export default SitePanel;
