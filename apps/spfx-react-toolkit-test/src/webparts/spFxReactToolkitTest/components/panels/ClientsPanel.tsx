import { useDemoStyles } from '../shared/demoStyles';
import { Select, Button, Field, Input, MessageBar, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { AadHttpClient, HttpClient, SPHttpClient, } from '@microsoft/sp-http';
import { useSPFxAadHttpClient, useSPFxAadTokenProvider, useSPFxApiPermissionPrecheck, useSPFxHttpClient, useSPFxMSGraphClient, useSPFxSPHttpClient, } from '@apvee/spfx-react-toolkit';
import { DemoCard, InfoGrid, JsonDetails, StatusBadge, } from '../shared';
/** Keep useful response fields visible while the complete payload stays in diagnostics. */
const responseRows = (value: unknown): Array<{
    label: string;
    value: string;
}> => {
    if (value === undefined)
        return [];
    if (value && typeof value === 'object') {
        return Object.entries(value).reduce<Array<{
            label: string;
            value: string;
        }>>((rows, [label, field]: [
            string,
            unknown
        ]) => {
            if (typeof field === 'string' || typeof field === 'number' || typeof field === 'boolean')
                rows.push({ label, value: String(field) });
            return rows;
        }, []).slice(0, 8);
    }
    return [{ label: 'Response', value: String(value) }];
};
const ClientsPanelStatusScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const http = useSPFxHttpClient();
    const spHttp = useSPFxSPHttpClient();
    const graph = useSPFxMSGraphClient();
    const aad = useSPFxAadHttpClient();
    const aadTokenProvider = useSPFxAadTokenProvider();
    return <DemoCard controls={selector(false)} title="Client Status" description="Inspect the availability and initialization of the SPFx clients." iconName="Cloud" contentPlacement="configuration">
      <div className={demoStyles.actions}>
        <StatusBadge label="HttpClient" available={http.isReady}/>
        <StatusBadge label="SPHttpClient" available={spHttp.isReady}/>
        <StatusBadge label="MSGraphClient" available={graph.isReady}/>
        <StatusBadge label="AadHttpClient" available={aad.isReady}/>
        <StatusBadge label="AadTokenProvider" available={aadTokenProvider.isReady}/>
      </div>
      <InfoGrid rows={[
            { label: 'SPHttp Base URL', value: spHttp.baseUrl, icon: 'Link' },
            { label: 'Graph Initializing', value: graph.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
            { label: 'AAD Resource', value: aad.resourceUrl, icon: 'CloudSecure' },
            { label: 'AAD Initializing', value: aad.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
            { label: 'AAD Token Provider Initializing', value: aadTokenProvider.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
        ]}/>
    </DemoCard>;
};
const ClientsPanelHttpScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const http = useSPFxHttpClient();
    const [httpResult, setHttpResult] = React.useState<unknown>();
    React.useLayoutEffect(() => { setHttpResult(undefined); }, [http.invoke]);
    const latest = React.useRef({ http: http.invoke });
    latest.current = { http: http.invoke };
    const loadExternal = React.useCallback(async () => {
        const invoked = latest.current.http;
        setHttpResult(undefined);
        const data = await http.invoke(client => client
            .get('https://jsonplaceholder.typicode.com/todos/1', HttpClient.configurations.v1)
            .then(response => response.json()));
        if (mounted.current && latest.current.http === invoked)
            setHttpResult(data);
    }, [http]);
    return <DemoCard controls={selector(false)} title="HttpClient" description="Load a sample todo from an external HTTP endpoint." iconName="CloudDownload" error={http.error} contentPlacement="configuration" commands={<>
      <Button appearance="primary" disabled={!http.isReady || http.isLoading} onClick={loadExternal}>{http.isLoading ? 'Loading...' : 'Load external todo'}</Button>
    </>} diagnostics={httpResult === undefined ? undefined : <JsonDetails label="HttpClient result" value={httpResult}/>}><InfoGrid rows={[{ label: 'Ready', value: http.isReady ? 'Yes' : 'No' }, { label: 'Response', value: httpResult === undefined ? 'Not loaded' : 'Loaded' }, ...responseRows(http.isLoading || http.error ? undefined : httpResult)]}/></DemoCard>;
};
const ClientsPanelSharepointScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const spHttp = useSPFxSPHttpClient();
    const [spResult, setSpResult] = React.useState<unknown>();
    React.useLayoutEffect(() => { setSpResult(undefined); }, [spHttp.invoke]);
    const latest = React.useRef({ sp: spHttp.invoke });
    latest.current = { sp: spHttp.invoke };
    const loadCurrentWeb = React.useCallback(async () => {
        const invoked = latest.current.sp;
        setSpResult(undefined);
        const data = await spHttp.invoke(client => client
            .get(`${spHttp.baseUrl}/_api/web?$select=Title,Url`, SPHttpClient.configurations.v1)
            .then(response => response.json()));
        if (mounted.current && latest.current.sp === invoked)
            setSpResult(data);
    }, [spHttp]);
    return <DemoCard controls={selector(false)} title="SPHttpClient" description="Read the title and URL of the current SharePoint web." iconName="SharepointLogo" error={spHttp.error} contentPlacement="configuration" commands={<>
      <Button appearance="primary" disabled={!spHttp.isReady || spHttp.isLoading} onClick={loadCurrentWeb}>{spHttp.isLoading ? 'Loading...' : 'Load current web'}</Button>
    </>} diagnostics={spResult === undefined ? undefined : <JsonDetails label="SPHttpClient result" value={spResult}/>}><InfoGrid rows={[{ label: 'Ready', value: spHttp.isReady ? 'Yes' : 'No' }, { label: 'Response', value: spResult === undefined ? 'Not loaded' : 'Loaded' }, ...responseRows(spHttp.isLoading || spHttp.error ? undefined : spResult)]}/></DemoCard>;
};
const ClientsPanelGraphScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const graph = useSPFxMSGraphClient();
    const [graphResult, setGraphResult] = React.useState<unknown>();
    React.useLayoutEffect(() => { setGraphResult(undefined); }, [graph.invoke]);
    const latest = React.useRef({ graph: graph.invoke });
    latest.current = { graph: graph.invoke };
    const loadGraphMe = React.useCallback(async () => {
        const invoked = latest.current.graph;
        setGraphResult(undefined);
        const data = await graph.invoke(client => client.api('/me').select('displayName,mail,userPrincipalName').get());
        if (mounted.current && latest.current.graph === invoked)
            setGraphResult(data);
    }, [graph]);
    return <DemoCard controls={selector(false)} title="MSGraphClient" description="Read the current user profile from Microsoft Graph." iconName="OfficeAssistantLogo" error={graph.error ?? graph.initError} contentPlacement="configuration" commands={<>
      <Button appearance="primary" disabled={!graph.isReady || graph.isLoading} onClick={loadGraphMe}>{graph.isLoading ? 'Loading...' : 'Load /me'}</Button>
    </>} diagnostics={<>
      <MessageBar intent="info">
        <MessageBarBody>
          Requires Graph permissions granted to this package, typically User.Read.
        </MessageBarBody>
      </MessageBar>
      {graphResult && <JsonDetails label="MSGraphClient result" value={graphResult}/>}
    </>}><InfoGrid rows={[{ label: 'Ready', value: graph.isReady ? 'Yes' : 'No' }, { label: 'Response', value: graphResult === undefined ? 'Not loaded' : 'Loaded' }, ...responseRows(graph.isLoading || graph.error ? undefined : graphResult)]}/></DemoCard>;
};
const ClientsPanelAadScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const aad = useSPFxAadHttpClient();
    const [aadResult, setAadResult] = React.useState<unknown>();
    const [aadResource, setAadResource] = React.useState('');
    const [aadPath, setAadPath] = React.useState('/api/health');
    const trimmedAadResource = React.useMemo(() => aadResource.trim(), [aadResource]);
    React.useLayoutEffect(() => { setAadResult(undefined); }, [aad.invoke, aad.resourceUrl, aadPath, trimmedAadResource]);
    const initializeAad = React.useCallback(() => {
        aad.setResourceUrl(trimmedAadResource);
        setAadResult(undefined);
    }, [aad, trimmedAadResource]);
    const latest = React.useRef({ aad: aad.invoke, target: `${trimmedAadResource}:${aad.resourceUrl}:${aadPath}` });
    latest.current = { aad: aad.invoke, target: `${trimmedAadResource}:${aad.resourceUrl}:${aadPath}` };
    const loadAad = React.useCallback(async () => {
        if (!aad.resourceUrl) {
            return;
        }
        const normalizedPath = aadPath.indexOf('/') === 0 ? aadPath : `/${aadPath}`;
        const target = latest.current.target;
        const invoked = latest.current.aad;
        setAadResult(undefined);
        const data = await aad.invoke(client => client
            .get(`${aad.resourceUrl}${normalizedPath}`, AadHttpClient.configurations.v1)
            .then(response => response.json()));
        if (mounted.current && latest.current.aad === invoked && latest.current.target === target)
            setAadResult(data);
    }, [aad, aadPath]);
    return <DemoCard title="AadHttpClient" description="Initialize a client for a custom API and call its health endpoint." iconName="AzureAPIManagement" error={aad.error ?? aad.initError} controls={<>{selector(false)}<>
      <div className={demoStyles.actions}>
        <Field label={"Resource URL or App ID URI"}>
          <Input value={aadResource} onChange={(_, data) => setAadResource(data.value ?? '')} placeholder="api://00000000-0000-0000-0000-000000000000"/>
        </Field>
        <Field label={"Path"}>
          <Input value={aadPath} onChange={(_, data) => setAadPath(data.value ?? '')}/>
        </Field>
      </div>

    </></>} contentPlacement="configuration" commands={<>
      <Button appearance="primary" disabled={!trimmedAadResource || aad.isInitializing} onClick={initializeAad}>{"Initialize AAD client"}</Button>
      <Button disabled={!aad.isReady || aad.isLoading} onClick={loadAad}>{aad.isLoading ? 'Loading...' : 'Call path'}</Button>
    </>} diagnostics={aadResult === undefined ? undefined : <JsonDetails label="AadHttpClient result" value={aadResult}/>}><InfoGrid rows={[{ label: 'Ready', value: aad.isReady ? 'Yes' : 'No' }, { label: 'Response', value: aadResult === undefined ? 'Not loaded' : 'Loaded' }, ...responseRows(aad.isLoading || aad.error ? undefined : aadResult)]}/></DemoCard>;
};
const ClientsPanelPermissionsScenario: React.FC<{
    readonly selector: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const aadTokenProvider = useSPFxAadTokenProvider();
    const [aadResource, setAadResource] = React.useState('');
    const trimmedAadResource = React.useMemo(() => aadResource.trim(), [aadResource]);
    const permissionPrecheckConfig = React.useMemo(() => ({
        graph: ['User.Read'],
        customApis: trimmedAadResource
            ? [
                {
                    name: 'Custom API',
                    resource: trimmedAadResource,
                    scopes: ['user_impersonation']
                }
            ]
            : []
    }), [trimmedAadResource]);
    const permissionPrecheckOptions = React.useMemo(() => ({
        autoCheck: false,
        mode: 'passive'
    } as const), []);
    const permissionPrecheck = useSPFxApiPermissionPrecheck(permissionPrecheckConfig, permissionPrecheckOptions);
    return <DemoCard controls={<>{selector(false)}<Field label="Resource URL or App ID URI"><Input value={aadResource} onChange={(_, data) => setAadResource(data.value)}/></Field></>} title="API Permission Precheck" description="Check configured Graph and custom API permissions." iconName="Permissions" error={permissionPrecheck.tokenProviderError} contentPlacement="configuration" commands={<>
      <Button appearance="primary" disabled={!aadTokenProvider.isReady || permissionPrecheck.isChecking} onClick={permissionPrecheck.check}>{permissionPrecheck.isChecking ? 'Checking...' : 'Check permissions'}</Button>
      <Button disabled={!aadTokenProvider.isReady || permissionPrecheck.isChecking} onClick={permissionPrecheck.retryWithoutCache}>{"Retry without cache"}</Button>
    </>} diagnostics={permissionPrecheck.results.length > 0 ? <JsonDetails label="API permission precheck results" value={permissionPrecheck.results}/> : undefined}>
      <InfoGrid rows={[
            { label: 'Configuration State', value: permissionPrecheck.configurationState, icon: 'StatusCircleQuestionMark' },
            { label: 'Configured', value: permissionPrecheck.isConfigured ? 'Yes' : 'No', icon: 'Completed' },
            { label: 'Required Graph Scope', value: 'User.Read', icon: 'OfficeAssistantLogo' },
            { label: 'Custom API Resource', value: trimmedAadResource || 'Not configured', icon: 'AzureAPIManagement' },
        ]}/>
      {permissionPrecheck.missing.map(permission => (<MessageBar intent="error" key={permission.id}>
        <MessageBarBody>
          {permission.message}
        </MessageBarBody>
      </MessageBar>))}
      {permissionPrecheck.warnings.map(permission => (<MessageBar intent="warning" key={permission.id}>
        <MessageBarBody>
          {permission.message}
        </MessageBarBody>
      </MessageBar>))}

    </DemoCard>;
};
const ClientsPanel: React.FC = () => {
    const demoStyles = useDemoStyles();
    const [scenario, setScenario] = React.useState('http');
    const selector = (busy = false): React.ReactNode => <Field className={demoStyles.controlRow} label="Operation"><Select disabled={busy} select={{ 'data-demo-variant': 'ClientsPanel' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={scenario} onChange={(_, data) => setScenario(data.value)}><option value="status">Client status</option><option value="http">External HTTP</option><option value="sharepoint">SharePoint</option><option value="graph">Microsoft Graph</option><option value="aad">Custom API</option><option value="permissions">API permissions</option></Select></Field>;
    return <>{scenario === 'status' && <ClientsPanelStatusScenario selector={selector}/>}{scenario === 'http' && <ClientsPanelHttpScenario selector={selector}/>}{scenario === 'sharepoint' && <ClientsPanelSharepointScenario selector={selector}/>}{scenario === 'graph' && <ClientsPanelGraphScenario selector={selector}/>}{scenario === 'aad' && <ClientsPanelAadScenario selector={selector}/>}{scenario === 'permissions' && <ClientsPanelPermissionsScenario selector={selector}/>}</>;
};
export default ClientsPanel;
