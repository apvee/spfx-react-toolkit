import * as React from 'react';
import { Label, MessageBar, PrimaryButton, Stack, TextField } from '@fluentui/react';
import {
  useSPFxPnPListById,
  useSPFxPnPListByUrl,
  useSPFxPnPListByPath,
} from '@apvee/spfx-react-toolkit';
import type { SPFxPnPListInfo } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

interface DemoItem {
  readonly Id: number;
  // SharePoint returns null for an unset optional text column.
  // eslint-disable-next-line @rushstack/no-new-null
  readonly Title?: string | null;
}

interface ListQueryControlsProps {
  readonly title: string;
  readonly label: string;
  readonly action: string;
  readonly example: string;
  readonly value: string;
  readonly onChange: (value: string) => void;
  readonly list: SPFxPnPListInfo<DemoItem>;
}

const ListQueryControls: React.FC<ListQueryControlsProps> = ({ title, label, action, example, value, onChange, list }) => {
  const handleQuery = React.useCallback(() => {
    // The hook publishes failures through error; consume the rejected action promise.
    list.query(q => q.select('Id', 'Title').orderBy('Id', false)).catch(() => undefined);
  }, [list.query]);

  return (
    <DemoCard title={title} iconName="BulletedList" error={list.error} description="Read the first 10 items, ordered by descending item ID. Queries run only when requested.">
      <TextField
        label={label}
        value={value}
        onChange={(_, nextValue) => onChange(nextValue ?? '')}
        placeholder={example}
      />
      <Label>Example: {example}</Label>
      <PrimaryButton text={action} onClick={handleQuery} disabled={!value.trim() || list.loading} />
      {list.loading && <MessageBar>Loading...</MessageBar>}
      {value.trim() && list.isEmpty && <MessageBar>No items loaded.</MessageBar>}
      {list.items.length > 0 && (
        <Stack tokens={{ childrenGap: 4 }}>
          <Label>Items ({list.items.length}):</Label>
          {list.items.map(item => <div key={item.Id}>#{item.Id} - {item.Title || '(untitled)'}</div>)}
        </Stack>
      )}
    </DemoCard>
  );
};

const ListByIdDemo: React.FC = () => {
  const [value, setValue] = React.useState('');
  const list = useSPFxPnPListById<DemoItem>(value, { pageSize: 10 });
  return <ListQueryControls title="useSPFxPnPListById" label="List GUID" action="Query by GUID" example="11111111-2222-3333-4444-555555555555" value={value} onChange={setValue} list={list} />;
};

const ListByUrlDemo: React.FC = () => {
  const [value, setValue] = React.useState('');
  const list = useSPFxPnPListByUrl<DemoItem>(value, { pageSize: 10 });
  return <ListQueryControls title="useSPFxPnPListByUrl" label="Server-relative list URL" action="Query by URL" example="/sites/projects/Lists/Tasks" value={value} onChange={setValue} list={list} />;
};

const ListByPathDemo: React.FC = () => {
  const [value, setValue] = React.useState('');
  const list = useSPFxPnPListByPath<DemoItem>(value, { pageSize: 10 });
  return <ListQueryControls title="useSPFxPnPListByPath" label="Web-relative list path" action="Query by path" example="Lists/Tasks" value={value} onChange={setValue} list={list} />;
};

export const PnPListAccessDemo: React.FC = () => (
  <Stack tokens={{ childrenGap: 16 }}>
    <MessageBar>Use the list GUID or decoded list root URL/path. URL and path examples use spaces and special characters as supplied. Paths are relative to the configured PnP web.</MessageBar>
    <ListByIdDemo />
    <ListByUrlDemo />
    <ListByPathDemo />
  </Stack>
);
