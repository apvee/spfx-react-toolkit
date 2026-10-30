import { useDemoStyles } from '../shared/demoStyles';
import { Select, Button, Field, Input, Label, MessageBar, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPListById, useSPFxPnPListByUrl, useSPFxPnPListByPath, } from '@apvee/spfx-react-toolkit';
import type { SPFxPnPListInfo } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
interface DemoItem {
    readonly Id: number;
    // SharePoint returns null for an unset optional text column.
    // eslint-disable-next-line @rushstack/no-new-null
    readonly Title?: string | null;
}
interface ListQueryControlsProps {
    readonly selector?: React.ReactNode;
    readonly title: string;
    readonly label: string;
    readonly action: string;
    readonly example: string;
    readonly value: string;
    readonly onChange: (value: string) => void;
    readonly list: SPFxPnPListInfo<DemoItem>;
}
const ListQueryControls: React.FC<ListQueryControlsProps> = ({ selector, title, label, action, example, value, onChange, list }) => {
    const demoStyles = useDemoStyles();
    const handleQuery = React.useCallback(() => {
        // The hook publishes failures through error; consume the rejected action promise.
        list.query(q => q.select('Id', 'Title').orderBy('Id', false)).catch(() => undefined);
    }, [list.query]);
    return (<DemoCard title={title} iconName="BulletedList" error={list.error} description="Read the first 10 items, ordered by descending item ID. Queries run only when requested." controls={<>{selector}
    <Field label={label}>
      <Input value={value} onChange={(_, data) => onChange(data.value ?? '')} placeholder={example}/>
    </Field>

  </>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={handleQuery} disabled={!value.trim() || list.loading}>{action}</Button>
  </>}>

    <Label>Example: {example}</Label>
    {list.loading && <MessageBar>
      <MessageBarBody>Loading...</MessageBarBody>
    </MessageBar>}
    {value.trim() && !list.loading && !list.error && list.isEmpty && <MessageBar>
      <MessageBarBody>No items loaded.</MessageBarBody>
    </MessageBar>}
    {!list.loading && !list.error && list.items.length > 0 && (<div className={demoStyles.panel}>
      <Label>Items ({list.items.length}):</Label>
      {list.items.map(item => <div key={item.Id}>#{item.Id} - {item.Title || '(untitled)'}</div>)}
    </div>)}
  </DemoCard>);
};
const ListByIdDemo: React.FC<{
    readonly selector: React.ReactNode;
}> = ({ selector }) => {
    const [value, setValue] = React.useState('');
    const list = useSPFxPnPListById<DemoItem>(value, { pageSize: 10 });
    return <ListQueryControls selector={selector} title="useSPFxPnPListById" label="List GUID" action="Query by GUID" example="11111111-2222-3333-4444-555555555555" value={value} onChange={setValue} list={list}/>;
};
const ListByUrlDemo: React.FC<{
    readonly selector: React.ReactNode;
}> = ({ selector }) => {
    const [value, setValue] = React.useState('');
    const list = useSPFxPnPListByUrl<DemoItem>(value, { pageSize: 10 });
    return <ListQueryControls selector={selector} title="useSPFxPnPListByUrl" label="Server-relative list URL" action="Query by URL" example="/sites/projects/Lists/Tasks" value={value} onChange={setValue} list={list}/>;
};
const ListByPathDemo: React.FC<{
    readonly selector: React.ReactNode;
}> = ({ selector }) => {
    const [value, setValue] = React.useState('');
    const list = useSPFxPnPListByPath<DemoItem>(value, { pageSize: 10 });
    return <ListQueryControls selector={selector} title="useSPFxPnPListByPath" label="Web-relative list path" action="Query by path" example="Lists/Tasks" value={value} onChange={setValue} list={list}/>;
};
export const PnPListAccessDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const [method, setMethod] = React.useState('id');
    const controls = <>{selector?.(false)}<Field label="List access"><Select select={{ 'data-demo-variant': 'PnPListAccessDemo' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={method} onChange={(_, data) => setMethod(data.value)}>
    <option value="id">GUID</option><option value="url">URL</option><option value="path">Path</option>
  </Select></Field></>;
    return method === 'id' ? <ListByIdDemo selector={controls}/> : method === 'url' ? <ListByUrlDemo selector={controls}/> : <ListByPathDemo selector={controls}/>;
};
