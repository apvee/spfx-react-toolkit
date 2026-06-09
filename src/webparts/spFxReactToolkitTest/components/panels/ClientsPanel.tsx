import * as React from 'react';
import {
  AadHttpClient,
  HttpClient,
  SPHttpClient,
} from '@microsoft/sp-http';
import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack,
  TextField,
} from '@fluentui/react';
import {
  useSPFxAadHttpClient,
  useSPFxAadTokenProvider,
  useSPFxApiPermissionPrecheck,
  useSPFxHttpClient,
  useSPFxMSGraphClient,
  useSPFxSPHttpClient,
} from '../../../../hooks';
import {
  DemoCard,
  InfoGrid,
  JsonDetails,
  StatusBadge,
} from '../shared';

const ClientsPanel: React.FC = () => {
  const http = useSPFxHttpClient();
  const spHttp = useSPFxSPHttpClient();
  const graph = useSPFxMSGraphClient();
  const aad = useSPFxAadHttpClient();
  const aadTokenProvider = useSPFxAadTokenProvider();

  const [httpResult, setHttpResult] = React.useState<unknown>();
  const [spResult, setSpResult] = React.useState<unknown>();
  const [graphResult, setGraphResult] = React.useState<unknown>();
  const [aadResult, setAadResult] = React.useState<unknown>();
  const [aadResource, setAadResource] = React.useState('');
  const [aadPath, setAadPath] = React.useState('/api/health');
  const trimmedAadResource = React.useMemo(
    () => aadResource.trim(),
    [aadResource]
  );
  const permissionPrecheckConfig = React.useMemo(
    () => ({
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
    }),
    [trimmedAadResource]
  );
  const permissionPrecheckOptions = React.useMemo(
    () => ({
      autoCheck: false,
      mode: 'passive'
    } as const),
    []
  );
  const permissionPrecheck = useSPFxApiPermissionPrecheck(
    permissionPrecheckConfig,
    permissionPrecheckOptions
  );

  const loadExternal = React.useCallback(async () => {
    const data = await http.invoke(client =>
      client
        .get('https://jsonplaceholder.typicode.com/todos/1', HttpClient.configurations.v1)
        .then(response => response.json())
    );
    setHttpResult(data);
  }, [http]);

  const loadCurrentWeb = React.useCallback(async () => {
    const data = await spHttp.invoke(client =>
      client
        .get(`${spHttp.baseUrl}/_api/web?$select=Title,Url`, SPHttpClient.configurations.v1)
        .then(response => response.json())
    );
    setSpResult(data);
  }, [spHttp]);

  const loadGraphMe = React.useCallback(async () => {
    const data = await graph.invoke(client =>
      client.api('/me').select('displayName,mail,userPrincipalName').get()
    );
    setGraphResult(data);
  }, [graph]);

  const initializeAad = React.useCallback(() => {
    aad.setResourceUrl(trimmedAadResource);
    setAadResult(undefined);
  }, [aad, trimmedAadResource]);

  const loadAad = React.useCallback(async () => {
    if (!aad.resourceUrl) {
      return;
    }

    const normalizedPath = aadPath.indexOf('/') === 0 ? aadPath : `/${aadPath}`;
    const data = await aad.invoke(client =>
      client
        .get(`${aad.resourceUrl}${normalizedPath}`, AadHttpClient.configurations.v1)
        .then(response => response.json())
    );
    setAadResult(data);
  }, [aad, aadPath]);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <DemoCard title="Client Status" iconName="Cloud">
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <StatusBadge label="HttpClient" available={http.isReady} />
          <StatusBadge label="SPHttpClient" available={spHttp.isReady} />
          <StatusBadge label="MSGraphClient" available={graph.isReady} />
          <StatusBadge label="AadHttpClient" available={aad.isReady} />
          <StatusBadge label="AadTokenProvider" available={aadTokenProvider.isReady} />
        </Stack>
        <InfoGrid
          rows={[
            { label: 'SPHttp Base URL', value: spHttp.baseUrl, icon: 'Link' },
            { label: 'Graph Initializing', value: graph.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
            { label: 'AAD Resource', value: aad.resourceUrl, icon: 'CloudSecure' },
            { label: 'AAD Initializing', value: aad.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
            { label: 'AAD Token Provider Initializing', value: aadTokenProvider.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
          ]}
        />
      </DemoCard>

      <DemoCard title="HttpClient" iconName="CloudDownload" error={http.error}>
        <PrimaryButton
          text={http.isLoading ? 'Loading...' : 'Load external todo'}
          iconProps={{ iconName: 'Download' }}
          disabled={!http.isReady || http.isLoading}
          onClick={loadExternal}
        />
        {httpResult && <JsonDetails label="HttpClient result" value={httpResult} />}
      </DemoCard>

      <DemoCard title="SPHttpClient" iconName="SharepointLogo" error={spHttp.error}>
        <PrimaryButton
          text={spHttp.isLoading ? 'Loading...' : 'Load current web'}
          iconProps={{ iconName: 'Download' }}
          disabled={!spHttp.isReady || spHttp.isLoading}
          onClick={loadCurrentWeb}
        />
        {spResult && <JsonDetails label="SPHttpClient result" value={spResult} />}
      </DemoCard>

      <DemoCard title="MSGraphClient" iconName="OfficeAssistantLogo" error={graph.error ?? graph.initError}>
        <PrimaryButton
          text={graph.isLoading ? 'Loading...' : 'Load /me'}
          iconProps={{ iconName: 'Contact' }}
          disabled={!graph.isReady || graph.isLoading}
          onClick={loadGraphMe}
        />
        <MessageBar messageBarType={MessageBarType.info}>
          Requires Graph permissions granted to this package, typically User.Read.
        </MessageBar>
        {graphResult && <JsonDetails label="MSGraphClient result" value={graphResult} />}
      </DemoCard>

      <DemoCard title="AadHttpClient" iconName="AzureAPIManagement" error={aad.error ?? aad.initError}>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <TextField
            label="Resource URL or App ID URI"
            value={aadResource}
            onChange={(_, value) => setAadResource(value ?? '')}
            placeholder="api://00000000-0000-0000-0000-000000000000"
            styles={{ root: { minWidth: 320 } }}
          />
          <TextField
            label="Path"
            value={aadPath}
            onChange={(_, value) => setAadPath(value ?? '')}
            styles={{ root: { minWidth: 180 } }}
          />
        </Stack>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Initialize AAD client"
            iconProps={{ iconName: 'PlugConnected' }}
            disabled={!trimmedAadResource || aad.isInitializing}
            onClick={initializeAad}
          />
          <DefaultButton
            text={aad.isLoading ? 'Loading...' : 'Call path'}
            iconProps={{ iconName: 'Send' }}
            disabled={!aad.isReady || aad.isLoading}
            onClick={loadAad}
          />
        </Stack>
        {aadResult && <JsonDetails label="AadHttpClient result" value={aadResult} />}
      </DemoCard>

      <DemoCard title="API Permission Precheck" iconName="Permissions" error={permissionPrecheck.tokenProviderError}>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={permissionPrecheck.isChecking ? 'Checking...' : 'Check permissions'}
            iconProps={{ iconName: 'Permissions' }}
            disabled={!aadTokenProvider.isReady || permissionPrecheck.isChecking}
            onClick={permissionPrecheck.check}
          />
          <DefaultButton
            text="Retry without cache"
            iconProps={{ iconName: 'Refresh' }}
            disabled={!aadTokenProvider.isReady || permissionPrecheck.isChecking}
            onClick={permissionPrecheck.retryWithoutCache}
          />
        </Stack>
        <InfoGrid
          rows={[
            { label: 'Configuration State', value: permissionPrecheck.configurationState, icon: 'StatusCircleQuestionMark' },
            { label: 'Configured', value: permissionPrecheck.isConfigured ? 'Yes' : 'No', icon: 'Completed' },
            { label: 'Required Graph Scope', value: 'User.Read', icon: 'OfficeAssistantLogo' },
            { label: 'Custom API Resource', value: trimmedAadResource || 'Not configured', icon: 'AzureAPIManagement' },
          ]}
        />
        {permissionPrecheck.missing.map(permission => (
          <MessageBar key={permission.id} messageBarType={MessageBarType.error}>
            {permission.message}
          </MessageBar>
        ))}
        {permissionPrecheck.warnings.map(permission => (
          <MessageBar key={permission.id} messageBarType={MessageBarType.warning}>
            {permission.message}
          </MessageBar>
        ))}
        {permissionPrecheck.results.length > 0 && (
          <JsonDetails label="API permission precheck results" value={permissionPrecheck.results} />
        )}
      </DemoCard>
    </Stack>
  );
};

export default ClientsPanel;
