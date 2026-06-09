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

  const [httpResult, setHttpResult] = React.useState<unknown>();
  const [spResult, setSpResult] = React.useState<unknown>();
  const [graphResult, setGraphResult] = React.useState<unknown>();
  const [aadResult, setAadResult] = React.useState<unknown>();
  const [aadResource, setAadResource] = React.useState('');
  const [aadPath, setAadPath] = React.useState('/api/health');

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
    aad.setResourceUrl(aadResource.trim());
    setAadResult(undefined);
  }, [aad, aadResource]);

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
        </Stack>
        <InfoGrid
          rows={[
            { label: 'SPHttp Base URL', value: spHttp.baseUrl, icon: 'Link' },
            { label: 'Graph Initializing', value: graph.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
            { label: 'AAD Resource', value: aad.resourceUrl, icon: 'CloudSecure' },
            { label: 'AAD Initializing', value: aad.isInitializing ? 'Yes' : 'No', icon: 'Sync' },
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
            disabled={!aadResource.trim() || aad.isInitializing}
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
    </Stack>
  );
};

export default ClientsPanel;
