import * as React from 'react';
import {
  DefaultButton,
  Image,
  ImageFit,
  Label,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack,
  TextField,
} from '@fluentui/react';
import {
  useSPFxOneDriveAppData,
  useSPFxUserInfo,
  useSPFxUserPhoto,
} from '../../../../hooks';
import {
  DemoCard,
  InfoGrid,
  JsonDetails,
} from '../shared';

interface OneDriveDemoData {
  readonly message: string;
  readonly counter: number;
  readonly timestamp: number;
}

const defaultData: OneDriveDemoData = {
  message: '',
  counter: 0,
  timestamp: 0,
};

const GraphDataPanel: React.FC = () => {
  const user = useSPFxUserInfo();
  const photo = useSPFxUserPhoto({ autoFetch: false });
  const oneDriveOptions = React.useMemo(() => ({
    autoFetch: false,
    createIfMissing: true,
    defaultValue: defaultData,
    folder: 'spfx-react-toolkit-demo',
  }), []);
  const appData = useSPFxOneDriveAppData<OneDriveDemoData>('test-data.json', oneDriveOptions);
  const [message, setMessage] = React.useState('');

  const saveData = React.useCallback(async () => {
    await appData.write({
      message,
      counter: (appData.data?.counter ?? 0) + 1,
      timestamp: Date.now(),
    });
    setMessage('');
  }, [appData, message]);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <DemoCard title="User Photo" iconName="ContactCard" error={photo.error}>
        <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center">
          {photo.photoUrl ? (
            <Image
              src={photo.photoUrl}
              alt={user.displayName}
              width={96}
              height={96}
              imageFit={ImageFit.cover}
              styles={{ image: { borderRadius: '50%' } }}
            />
          ) : (
            <div style={{
              alignItems: 'center',
              background: '#0078d4',
              borderRadius: '50%',
              color: '#fff',
              display: 'flex',
              fontSize: 36,
              fontWeight: 600,
              height: 96,
              justifyContent: 'center',
              width: 96,
            }}>
              {user.displayName ? user.displayName.charAt(0).toUpperCase() : '?'}
            </div>
          )}
          <InfoGrid
            rows={[
              { label: 'Display Name', value: user.displayName, icon: 'Contact' },
              { label: 'Ready', value: photo.isReady ? 'Yes' : 'No', icon: 'CheckMark' },
              { label: 'Blob', value: photo.photoBlob ? `${photo.photoBlob.size} bytes` : undefined, icon: 'FileImage' },
            ]}
          />
        </Stack>
        <PrimaryButton
          text={photo.isLoading ? 'Loading...' : 'Load photo'}
          iconProps={{ iconName: 'Refresh' }}
          disabled={photo.isLoading}
          onClick={photo.reload}
        />
        <MessageBar messageBarType={MessageBarType.info}>
          Uses Microsoft Graph photo endpoints and cleans up generated blob URLs on unmount.
        </MessageBar>
      </DemoCard>

      <DemoCard title="OneDrive AppData" iconName="OneDrive" error={appData.error ?? appData.writeError}>
        <InfoGrid
          rows={[
            { label: 'Ready', value: appData.isReady ? 'Yes' : 'No', icon: 'CheckMark' },
            { label: 'Not Found', value: appData.isNotFound ? 'Yes' : 'No', icon: 'BlockedSiteSolid12' },
            { label: 'Loading', value: appData.isLoading ? 'Yes' : 'No', icon: 'Sync' },
            { label: 'Writing', value: appData.isWriting ? 'Yes' : 'No', icon: 'CloudUpload' },
          ]}
        />
        <TextField
          label="Message"
          value={message}
          onChange={(_, value) => setMessage(value ?? '')}
          placeholder="Message to store in OneDrive approot"
        />
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={appData.isLoading ? 'Loading...' : 'Load'}
            iconProps={{ iconName: 'CloudDownload' }}
            disabled={appData.isLoading || appData.isWriting}
            onClick={appData.load}
          />
          <DefaultButton
            text={appData.isWriting ? 'Saving...' : 'Write'}
            iconProps={{ iconName: 'CloudUpload' }}
            disabled={!message || appData.isLoading || appData.isWriting}
            onClick={saveData}
          />
        </Stack>
        {appData.data && (
          <>
            <Label>Data is stored in approot:/spfx-react-toolkit-demo/test-data.json.</Label>
            <JsonDetails label="OneDrive data" value={appData.data} />
          </>
        )}
      </DemoCard>
    </Stack>
  );
};

export default GraphDataPanel;
