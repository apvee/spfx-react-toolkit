import { useDemoStyles } from '../shared/demoStyles';
import { Avatar, Select, Button, Field, Image, Input, Label, MessageBar, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxOneDriveAppData, useSPFxMSGraphClient, useSPFxUserInfo, useSPFxUserPhoto, } from '@apvee/spfx-react-toolkit';
import { DemoCard, InfoGrid, JsonDetails, } from '../shared';
interface OneDriveDemoData {
    readonly message: string;
    readonly counter: number;
    readonly timestamp: number;
}
type Selector = (busy?: boolean) => React.ReactNode;
const UserPhotoScenario: React.FC<{
    readonly selector: Selector;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const user = useSPFxUserInfo();
    const photo = useSPFxUserPhoto({ autoFetch: false });
    return <DemoCard title="User Photo" description="Read the current profile photo." iconName="ContactCard" error={photo.error} contentPlacement="configuration" controls={selector()} commands={<Button appearance="primary" disabled={photo.isLoading} onClick={photo.reload}>Run</Button>} diagnostics={<InfoGrid rows={[{ label: 'Ready', value: String(photo.isReady) }, { label: 'Blob', value: photo.photoBlob ? `${photo.photoBlob.size} bytes` : undefined }]}/>}>
    {photo.isLoading ? <Label>Loading...</Label> : <div className={demoStyles.actions}>{photo.photoUrl ? <Image fit="cover" shape="circular" src={photo.photoUrl} alt={user.displayName} width={96} height={96}/> : <Avatar name={user.displayName} size={96}/>}<Label>{user.displayName}</Label></div>}
  </DemoCard>;
};
const OneDriveScenario: React.FC<{
    readonly selector: Selector;
}> = ({ selector }) => {
    const options = React.useMemo(() => ({ autoFetch: false, createIfMissing: false, folder: 'spfx-react-toolkit-demo' }), []);
    const graph = useSPFxMSGraphClient();
    const appData = useSPFxOneDriveAppData<OneDriveDemoData>('test-data.json', options);
    const [message, setMessage] = React.useState('Project update');
    const [operation, setOperation] = React.useState('read');
    const [executed, setExecuted] = React.useState<{
        readonly operation: string;
        readonly request: number;
        readonly completed: boolean;
        readonly load: typeof appData.load;
        readonly write: typeof appData.write;
    }>();
    const [failure, setFailure] = React.useState<Error>();
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const generation = React.useRef(0);
    const identity = React.useRef({ load: appData.load, write: appData.write });
    if (identity.current.load !== appData.load || identity.current.write !== appData.write) {
        identity.current = { load: appData.load, write: appData.write };
        generation.current += 1;
    }
    const available = graph.isReady && graph.client !== undefined;
    const execute = React.useCallback(async () => {
        if (!available) return;
        const snapshot = { operation, load: appData.load, write: appData.write, request: ++generation.current, completed: false };
        const isCurrent = (): boolean => mounted.current && generation.current === snapshot.request && identity.current.load === snapshot.load && identity.current.write === snapshot.write;
        setExecuted(snapshot);
        setFailure(undefined);
        try {
            if (operation === 'read') await snapshot.load();
            else {
                await snapshot.write({ message, counter: (appData.data?.counter ?? 0) + 1, timestamp: Date.now() });
                if (isCurrent()) setMessage('');
            }
            if (isCurrent()) setExecuted({ ...snapshot, completed: true });
        } catch (error) {
            if (isCurrent()) {
                setFailure(error instanceof Error ? error : new Error(String(error)));
                setExecuted({ ...snapshot, completed: true });
            }
        }
    }, [appData, operation, message, available]);
    const busy = appData.isLoading || appData.isWriting;
    const current = executed?.request === generation.current && executed.operation === operation && executed.load === appData.load && executed.write === appData.write;
    const error = current ? failure ?? (operation === 'read' ? appData.error : appData.writeError) : undefined;
    return <DemoCard title="OneDrive AppData" description="Read the saved message or write it with an explicit command." iconName="OneDrive" error={error} contentPlacement="configuration" controls={<>{selector(appData.isWriting)}<Field label="OneDrive operation"><Select disabled={appData.isWriting} select={{ 'data-demo-variant': 'OneDriveScenario' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={operation} onChange={(_, data) => { generation.current += 1; setOperation(data.value); setExecuted(undefined); }}><option value="read">Read</option><option value="write">Write</option></Select></Field>{operation === 'write' && <Field label="Message"><Input disabled={appData.isWriting} value={message} onChange={(_, data) => { generation.current += 1; setMessage(data.value); setExecuted(undefined); }}/></Field>}</>} commands={<Button appearance="primary" disabled={!available || busy || (operation === 'write' && !message)} onClick={execute}>Run</Button>} diagnostics={<><Label>approot:/spfx-react-toolkit-demo/test-data.json</Label><InfoGrid rows={[{ label: 'Ready', value: String(appData.isReady) }, { label: 'Not Found', value: String(appData.isNotFound) }]}/>{current && executed?.completed && !busy && !error && appData.data && <JsonDetails label="OneDrive data" value={appData.data}/>}<MessageBar intent="info"><MessageBarBody>Reading a missing file does not create it. Only Write modifies OneDrive.</MessageBarBody></MessageBar></>}>
    {!available ? <Label>Microsoft Graph unavailable. Wait for client initialization.</Label> : busy ? <Label>Loading...</Label> : !current || !executed?.completed ? <Label>Run the selected operation.</Label> : error ? <Label>Operation not completed.</Label> : appData.isNotFound ? <Label>File not found.</Label> : appData.data === undefined ? <Label>No data returned.</Label> : <InfoGrid rows={[{ label: 'Saved message', value: appData.data?.message || '(empty)' }, { label: 'Counter', value: appData.data ? String(appData.data.counter) : undefined }]}/>}
  </DemoCard>;
};
const GraphDataPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('photo');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'GraphDataPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}><option value="photo">User photo</option><option value="onedrive">OneDrive</option></Select></Field>;
    return scenario === 'photo' ? <UserPhotoScenario selector={selector}/> : <OneDriveScenario selector={selector}/>;
};
export default GraphDataPanel;
