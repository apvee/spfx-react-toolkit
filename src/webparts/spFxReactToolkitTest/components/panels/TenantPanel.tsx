import * as React from 'react';
import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack,
  TextField,
} from '@fluentui/react';
import {
  useSPFxTenantKeyValueStore,
  useSPFxTenantProperty,
} from '../../../../hooks';
import type { SPFxTenantKeyValueStoreItem } from '../../../../hooks';
import {
  ActionResult,
  DemoCard,
  InfoGrid,
  JsonDetails,
  StatusBadge,
} from '../shared';

const TenantPanel: React.FC = () => {
  const tenantVersion = useSPFxTenantProperty<string>('spfx-toolkit-test-version', false);
  const tenantCounter = useSPFxTenantProperty<number>('spfx-toolkit-test-counter', false);
  const store = useSPFxTenantKeyValueStore();

  const [key, setKey] = React.useState('spfx-toolkit-demo');
  const [value, setValue] = React.useState('');
  const [description, setDescription] = React.useState('Created by SPFx React Toolkit sample webpart');
  const [storeItems, setStoreItems] = React.useState<SPFxTenantKeyValueStoreItem<unknown>[]>([]);
  const [selectedItem, setSelectedItem] = React.useState<SPFxTenantKeyValueStoreItem<unknown>>();
  const [result, setResult] = React.useState<string>();

  const loadProperties = React.useCallback(async () => {
    await Promise.all([tenantVersion.load(), tenantCounter.load()]);
    setResult('Tenant properties loaded.');
  }, [tenantVersion, tenantCounter]);

  const listItems = React.useCallback(async () => {
    const items = await store.list();
    setStoreItems(items);
    setResult(`Loaded ${items.length} tenant key-value item(s).`);
  }, [store]);

  const getItem = React.useCallback(async () => {
    const item = await store.get<string>(key);
    setSelectedItem(item);
    setResult(item ? `Found ${key}.` : `${key} not found.`);
  }, [store, key]);

  const saveItem = React.useCallback(async () => {
    await store.save<string>(key, value, description || undefined);
    setResult(`Saved ${key}.`);
  }, [store, key, value, description]);

  const removeItem = React.useCallback(async () => {
    await store.remove(key);
    setSelectedItem(undefined);
    setResult(`Removed ${key}.`);
  }, [store, key]);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <ActionResult result={result} error={tenantVersion.error ?? tenantCounter.error ?? store.error ?? store.writeError} />

      <DemoCard title="Tenant Properties" iconName="ServerEnviroment">
        <MessageBar messageBarType={MessageBarType.info}>
          Tenant properties are read-only through REST. Use PowerShell for writes, or the key-value store below for REST CRUD.
        </MessageBar>
        <PrimaryButton
          text={tenantVersion.isLoading || tenantCounter.isLoading ? 'Loading...' : 'Load properties'}
          iconProps={{ iconName: 'CloudDownload' }}
          disabled={tenantVersion.isLoading || tenantCounter.isLoading}
          onClick={loadProperties}
        />
        <InfoGrid
          rows={[
            { label: 'spfx-toolkit-test-version', value: tenantVersion.data ?? '(not set)', icon: 'ServerEnviroment' },
            { label: 'Version Description', value: tenantVersion.description, icon: 'Info' },
            { label: 'spfx-toolkit-test-counter', value: tenantCounter.data === undefined ? '(not set)' : String(tenantCounter.data), icon: 'NumberField' },
            { label: 'Counter Description', value: tenantCounter.description, icon: 'Info' },
          ]}
        />
      </DemoCard>

      <DemoCard title="Tenant Key-Value Store" iconName="TableGroup">
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <StatusBadge label="Ready" available={store.isReady} />
          <StatusBadge label="Can Write" available={store.canWrite} />
          <StatusBadge label="Loading" available={!store.isLoading} />
          <StatusBadge label="Writing" available={!store.isWriting} />
        </Stack>

        <TextField
          label="Key"
          value={key}
          onChange={(_, newValue) => setKey(newValue ?? '')}
        />
        <TextField
          label="Value"
          value={value}
          onChange={(_, newValue) => setValue(newValue ?? '')}
        />
        <TextField
          label="Description"
          value={description}
          onChange={(_, newValue) => setDescription(newValue ?? '')}
        />

        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="List" iconProps={{ iconName: 'BulletedList' }} disabled={!store.isReady || store.isLoading} onClick={listItems} />
          <DefaultButton text="Get" iconProps={{ iconName: 'Search' }} disabled={!store.isReady || store.isLoading || !key} onClick={getItem} />
          <DefaultButton text="Save" iconProps={{ iconName: 'Save' }} disabled={!store.isReady || store.isWriting || !key || !value} onClick={saveItem} />
          <DefaultButton text="Remove" iconProps={{ iconName: 'Delete' }} disabled={!store.isReady || store.isWriting || !key} onClick={removeItem} />
        </Stack>

        {selectedItem && <JsonDetails label="Selected key-value item" value={selectedItem} />}
        {storeItems.length > 0 && <JsonDetails label="Tenant key-value items" value={storeItems} />}
      </DemoCard>
    </Stack>
  );
};

export default TenantPanel;
