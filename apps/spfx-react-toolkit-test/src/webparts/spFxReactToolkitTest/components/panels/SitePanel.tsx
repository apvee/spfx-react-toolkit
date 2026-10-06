import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { useSPFxPageContext, useSPFxSiteKeyValueStore } from '@apvee/spfx-react-toolkit';
import type { SPFxSiteKeyValueStoreItem } from '@apvee/spfx-react-toolkit';
import { ActionResult, DemoCard, InfoGrid, JsonDetails, StatusBadge } from '../shared';

type SiteAction = 'list' | 'get' | 'save' | 'remove';

const SitePanel: React.FC = () => {
  const pageContext = useSPFxPageContext();
  const store = useSPFxSiteKeyValueStore();
  const [key, setKey] = React.useState('spfx-toolkit-site-demo-disposable');
  const [value, setValue] = React.useState('');
  const [description, setDescription] = React.useState('Disposable SPFx React Toolkit site demo item');
  // get changes when the hook's collection or client changes, including in-place PageContext changes.
  const identity = store.get;
  const currentIdentity = React.useRef(identity);
  currentIdentity.current = identity;
  const mounted = React.useRef(true);
  const request = React.useRef(0);
  const [outcome, setOutcome] = React.useState<{
    identity: typeof identity;
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
    const isCurrent = (): boolean => mounted.current && currentIdentity.current === identity && request.current === operation;
    setOutcome({ identity, action });
    try {
      if (action === 'list') {
        const items = await store.list();
        if (isCurrent()) setOutcome({ identity, action, items, result: `Loaded ${items.length} site key-value item(s).` });
      } else if (action === 'get') {
        const selectedItem = await store.get<string>(key);
        if (isCurrent()) setOutcome({ identity, action, selectedItem, result: selectedItem ? `Found ${key}.` : `${key} not found.` });
      } else if (action === 'save') {
        await store.save<string>(key, value, description || undefined);
        if (isCurrent()) setOutcome({ identity, action, result: `Saved ${key}.` });
      } else {
        await store.remove(key);
        if (isCurrent()) setOutcome({ identity, action, result: `Removed ${key}.` });
      }
    } catch (failure) {
      if (isCurrent()) setOutcome({ identity, action, error: failure instanceof Error ? failure : new Error(String(failure)) });
    }
  }, [identity, store, key, value, description]);

  const currentOutcome = outcome?.identity === identity ? outcome : undefined;
  const isRead = currentOutcome?.action === 'get' || currentOutcome?.action === 'list';
  // Reads resolve fallback values on failure. Only current rendered error state can distinguish failure from absence.
  const error = currentOutcome?.error ?? (isRead ? store.error : store.writeError);
  const result = error ? undefined : currentOutcome?.result;
  const busy = store.isLoading || store.isWriting;

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <ActionResult result={result} error={error} />
      <DemoCard title="Site Key-Value Store" iconName="TableGroup">
        <MessageBar messageBarType={MessageBarType.info}>
          All subsites share SiteKeyValueStore in the collection root. Actions run only when clicked.
          Use disposable keys beginning with spfx-toolkit-site-demo- and remove them after testing.
          Saving may create the store; SharePoint enforces permissions. Can Write is an indicator.
        </MessageBar>
        <InfoGrid rows={[
          { label: 'Collection root URL', value: pageContext.site.absoluteUrl, icon: 'Info' },
          { label: 'Current web URL', value: pageContext.web.absoluteUrl, icon: 'Info' },
        ]} />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <StatusBadge label="Ready" available={store.isReady} />
          <StatusBadge label="Can Write" available={store.canWrite} />
          <StatusBadge label="Loading" available={!store.isLoading} />
          <StatusBadge label="Writing" available={!store.isWriting} />
        </Stack>
        <TextField label="Key" value={key} onChange={(_, next) => setKey(next ?? '')} />
        <TextField label="Value" value={value} onChange={(_, next) => setValue(next ?? '')} />
        <TextField label="Description" value={description} onChange={(_, next) => setDescription(next ?? '')} />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="List" iconProps={{ iconName: 'BulletedList' }} disabled={!store.isReady || busy} onClick={() => runAction('list')} />
          <DefaultButton text="Get" iconProps={{ iconName: 'Search' }} disabled={!store.isReady || busy || !key} onClick={() => runAction('get')} />
          <DefaultButton text="Save" iconProps={{ iconName: 'Save' }} disabled={!store.isReady || busy || !key} onClick={() => runAction('save')} />
          <DefaultButton text="Remove" iconProps={{ iconName: 'Delete' }} disabled={!store.isReady || busy || !key} onClick={() => runAction('remove')} />
        </Stack>
        {!error && currentOutcome?.selectedItem && <JsonDetails label="Selected site key-value item" value={currentOutcome.selectedItem} />}
        {!error && currentOutcome?.items && currentOutcome.items.length > 0 && <JsonDetails label="Site key-value items" value={currentOutcome.items} />}
      </DemoCard>
    </Stack>
  );
};

export default SitePanel;
