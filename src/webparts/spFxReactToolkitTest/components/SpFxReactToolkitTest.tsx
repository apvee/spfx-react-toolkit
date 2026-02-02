import * as React from 'react';
import styles from './SpFxReactToolkitTest.module.scss';
import {
  Stack,
  Pivot,
  PivotItem,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';
import { InfoRow, StatusBadge } from './shared';
import {
  HttpClientDemo,
  PnPContextDemo,
  PnPOperationsDemo,
  PnPListDemo,
  PnPSearchBasicDemo,
  PnPSearchAdvancedDemo,
  PnPSearchRefinersDemo,
  PnPSearchSuggestionsDemo,
} from './demos';

// Import hooks from SPFx React Toolkit (used in main component)
import {
  useSPFxProperties,
  useSPFxThemeInfo,
  useSPFxUserInfo,
  useSPFxEnvironmentInfo,
  useSPFxSiteInfo,
  useSPFxDisplayMode,
  useSPFxInstanceInfo,
  useSPFxPageContext,
  useSPFxTeams,
  useSPFxListInfo,
  useSPFxLocaleInfo,
  useSPFxHubSiteInfo,
  useSPFxCorrelationInfo,
  useSPFxPermissions,
  useSPFxCrossSitePermissions,
  useSPFxContainerSize,
  useSPFxContainerInfo,
  useSPFxSessionStorage,
  useSPFxLocalStorage,
  useSPFxLogger,
  useSPFxPageType,
  useSPFxServiceScope,
  useSPFxSPHttpClient,
  useSPFxMSGraphClient,
  useSPFxAadHttpClient,
  useSPFxPerformance,
  useSPFxFluent9ThemeInfo,
  useSPFxOneDriveAppData,
  useSPFxTenantProperty,
  useSPFxUserPhoto,
} from '../../../hooks';
import { SPPermission } from '@microsoft/sp-page-context';

interface IWebPartProps {
  description: string;
}

interface IOneDriveTestData {
  message: string;
  counter: number;
  timestamp: number;
}

// Helper function to safely stringify objects with circular references
const safeStringify = (obj: unknown, indent: number = 2): string => {
  const seen = new WeakSet();
  return JSON.stringify(obj, (key, value) => {
    // Skip private properties and service scopes to avoid circular references
    if (typeof key === 'string' && (key.indexOf('_') === 0 || key === 'serviceScope' || key === 'service')) {
      return '[Circular/Private]';
    }
    if (typeof value === 'object' && value !== null) {
      if (seen.has(value)) {
        return '[Circular Reference]';
      }
      seen.add(value);
    }
    return value;
  }, indent);
};

const SpFxReactToolkitTest: React.FC = () => {
  // Core Properties & Display
  const { properties, setProperties } = useSPFxProperties<IWebPartProps>();
  const { isEdit } = useSPFxDisplayMode();
  const { id, kind } = useSPFxInstanceInfo();

  // Theme & Environment
  const theme = useSPFxThemeInfo();
  const fluent9ThemeInfo = useSPFxFluent9ThemeInfo();
  const { type: envType } = useSPFxEnvironmentInfo();
  const isDarkTheme = theme?.isInverted ?? false;

  // User & Site Info
  const { displayName } = useSPFxUserInfo();
  const { title: siteTitle, webUrl, siteClassification } = useSPFxSiteInfo();
  const localeInfo = useSPFxLocaleInfo();

  // Teams Context
  const { supported: hasTeamsContext, theme: teamsTheme } = useSPFxTeams();

  // Page Context
  const pageContext = useSPFxPageContext();
  const pageTypeInfo = useSPFxPageType();

  // List & Hub Info (can be undefined)
  const listInfo = useSPFxListInfo();
  const hubInfo = useSPFxHubSiteInfo();

  // Container & Performance
  const containerSize = useSPFxContainerSize();
  const containerInfo = useSPFxContainerInfo();
  const performance = useSPFxPerformance();

  // Permissions
  const { hasWebPermission } = useSPFxPermissions();
  const canManageWeb = hasWebPermission(SPPermission.manageWeb);
  const canManageLists = hasWebPermission(SPPermission.manageLists);

  // Storage
  const sessionStorage = useSPFxSessionStorage('demo-session-key', '');
  const localStorage = useSPFxLocalStorage('demo-local-key', '');

  // Advanced
  const correlationInfo = useSPFxCorrelationInfo();
  const logger = useSPFxLogger();
  const serviceScope = useSPFxServiceScope();

  // HTTP Clients
  const spHttpClient = useSPFxSPHttpClient();
  const msGraphClient = useSPFxMSGraphClient();
  const aadHttpClient = useSPFxAadHttpClient();

  // OneDrive AppData hook (using hook name as folder for demo)
  const oneDriveData = useSPFxOneDriveAppData<IOneDriveTestData>(
    'test-data.json',
    { autoFetch: false, createIfMissing: true, defaultValue: { message: '', counter: 0, timestamp: 0 }, folder: 'useSPFxOneDriveAppData' }
  );

  // Tenant Property hook (tenant-wide configuration)
  const tenantVersion = useSPFxTenantProperty<string>('spfx-toolkit-test-version', false);
  const tenantCounter = useSPFxTenantProperty<number>('spfx-toolkit-test-counter', false);

  // User Photo hook (current user profile photo)
  const userPhoto = useSPFxUserPhoto();

  // Local state for interactive demos
  const [descriptionInput, setDescriptionInput] = React.useState(properties?.description ?? '');
  const [sessionStorageInput, setSessionStorageInput] = React.useState('');
  const [localStorageInput, setLocalStorageInput] = React.useState('');
  const [oneDriveMessage, setOneDriveMessage] = React.useState('');
  const [tenantVersionInput, setTenantVersionInput] = React.useState('');
  const [showMessage, setShowMessage] = React.useState(false);
  const [messageText, setMessageText] = React.useState('');
  const [performanceResult, setPerformanceResult] = React.useState<string>('');
  const [logMessages, setLogMessages] = React.useState<Array<{ level: string; message: string }>>([]);
  const [crossSiteUrl, setCrossSiteUrl] = React.useState<string | undefined>(undefined);

  // Cross-site permissions (fetch only when URL is set)
  const crossSitePermissions = useSPFxCrossSitePermissions(crossSiteUrl);

  // Sync descriptionInput when properties change
  React.useEffect(() => {
    setDescriptionInput(properties?.description ?? '');
  }, [properties?.description]);

  // Handlers
  const handleUpdateProperties = React.useCallback(() => {
    setProperties({ description: descriptionInput });
    setMessageText('Properties updated successfully!');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [descriptionInput, setProperties]);

  const handleSaveSessionStorage = React.useCallback(() => {
    sessionStorage.setValue(sessionStorageInput);
    setSessionStorageInput('');
    setMessageText('Value saved to session storage!');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [sessionStorageInput, sessionStorage]);

  const handleSaveLocalStorage = React.useCallback(() => {
    localStorage.setValue(localStorageInput);
    setLocalStorageInput('');
    setMessageText('Value saved to local storage!');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [localStorageInput, localStorage]);

  const handlePerformanceTest = React.useCallback(async () => {
    const result = await performance.time('demo-test', async () => {
      // Simulate async operation
      await new Promise(resolve => setTimeout(resolve, 500));
      return 'Test completed';
    });
    setPerformanceResult(`${result.result} in ${result.durationMs.toFixed(2)}ms`);
  }, [performance]);

  const handleLog = React.useCallback((level: 'info' | 'warning' | 'error', message: string) => {
    if (level === 'info') logger.info(message);
    else if (level === 'warning') logger.warn(message);
    else logger.error(message);

    setLogMessages(prev => [...prev.slice(-4), { level, message }]);
  }, [logger]);

  const handleLoadOneDrive = React.useCallback(async () => {
    await oneDriveData.load();
    setMessageText('OneDrive load completed. Check status below (missing file sets isNotFound=true).');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [oneDriveData]);

  const handleSaveOneDrive = React.useCallback(async () => {
    try {
      const newData: IOneDriveTestData = {
        message: oneDriveMessage,
        counter: (oneDriveData.data?.counter ?? 0) + 1,
        timestamp: Date.now(),
      };
      await oneDriveData.write(newData);
      setOneDriveMessage('');
      setMessageText('Data saved to OneDrive successfully!');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Save failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [oneDriveMessage, oneDriveData]);

  const handleLoadTenantProperty = React.useCallback(async () => {
    try {
      await Promise.all([
        tenantVersion.load(),
        tenantCounter.load()
      ]);
      setMessageText('Tenant properties loaded successfully!');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Load failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantVersion, tenantCounter]);

  const handleSaveTenantVersion = React.useCallback(async () => {
    if (!tenantVersion.canWrite) {
      setMessageText('Insufficient permissions to write tenant properties');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
      return;
    }

    try {
      await tenantVersion.write(
        tenantVersionInput,
        'Test version property from SPFx React Toolkit'
      );
      setTenantVersionInput('');
      setMessageText('Version saved to tenant properties!');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Save failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantVersionInput, tenantVersion]);

  const handleIncrementTenantCounter = React.useCallback(async () => {
    if (!tenantCounter.canWrite) {
      setMessageText('Insufficient permissions to write tenant properties');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
      return;
    }

    try {
      const newValue = (tenantCounter.data ?? 0) + 1;
      await tenantCounter.write(
        newValue,
        'Test counter property from SPFx React Toolkit'
      );
      setMessageText(`Counter incremented to ${newValue}!`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Increment failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantCounter]);

  const handleRemoveTenantProperty = React.useCallback(async (propertyName: 'version' | 'counter') => {
    const hook = propertyName === 'version' ? tenantVersion : tenantCounter;

    if (!hook.canWrite) {
      setMessageText('Insufficient permissions to remove tenant properties');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
      return;
    }

    if (!confirm(`Are you sure you want to remove the ${propertyName} property?`)) {
      return;
    }

    try {
      await hook.remove();
      setMessageText(`${propertyName} property removed!`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Remove failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantVersion, tenantCounter]);

  return (
    <section className={`${styles.spFxReactToolkitTest} ${hasTeamsContext ? styles.teams : ''}`}>
      {/* Message Bar */}
      {showMessage && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setShowMessage(false)}>
          {messageText}
        </MessageBar>
      )}

      {/* Main Content with Pivot Tabs */}
      <Pivot aria-label="SPFx React Toolkit Demo Tabs" style={{ marginTop: '20px' }}>

        {/* TAB 1: Overview */}
        <PivotItem headerText="Overview" itemIcon="Home">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: '16px' } }}>

            {/* Core Information Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Info" style={{ marginRight: '8px' }} />
                Core Information
                <Separator />
              </h3>
              <InfoRow label="Instance ID" value={id} icon="Fingerprint" />
              <InfoRow label="Component Kind" value={kind} icon="CubeShape" />
              <InfoRow label="Display Mode" value={isEdit ? 'Edit' : 'Read'} icon="Edit" />
              <InfoRow label="Environment Type" value={envType} icon="Globe" />
              <InfoRow label="Page Type" value={pageTypeInfo.pageType} icon="Page" />
            </Stack>

            {/* Properties Editor Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Settings" style={{ marginRight: '8px' }} />
                Properties Editor
                <Separator />
              </h3>
              <TextField
                label="Description Property"
                value={descriptionInput}
                onChange={(_, newValue) => setDescriptionInput(newValue ?? '')}
                placeholder="Enter description..."
              />
              <PrimaryButton
                text="Update Properties"
                onClick={handleUpdateProperties}
                iconProps={{ iconName: 'Save' }}
              />
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Changes will update the WebPart properties and trigger Property Pane refresh
              </Label>
            </Stack>

            {/* Container Info */}
            {containerSize && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="Resize" style={{ marginRight: '8px' }} />
                  Container Information
                  <Separator />
                </h3>
                <InfoRow label="Width" value={`${containerSize.width}px`} icon="ArrowRight" />
                <InfoRow label="Height" value={`${containerSize.height}px`} icon="ArrowUp" />
                <InfoRow label="DOM Element" value={containerInfo.element ? 'Available' : 'N/A'} icon="DOM" />
                <InfoRow label="Size Tracking" value={containerInfo.size ? 'Active' : 'Inactive'} icon="RadioBullet" />
              </Stack>
            )}

          </Stack>
        </PivotItem>

        {/* TAB 2: Context */}
        <PivotItem headerText="Context" itemIcon="Contact">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: '16px' } }}>

            {/* User & Site Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="People" style={{ marginRight: '8px' }} />
                User & Site Information
                <Separator />
              </h3>

              <InfoRow label="User Name" value={displayName} icon="Contact" />
              <InfoRow label="Site Title" value={siteTitle} icon="CityNext" />
              <InfoRow label="Site URL" value={webUrl} icon="Link" />
              {siteClassification && (
                <InfoRow label="Classification" value={siteClassification} icon="Tag" />
              )}
            </Stack>

            {/* Locale & Regional Settings */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="LocaleLanguage" style={{ marginRight: '8px' }} />
                Locale & Regional Settings
                <Separator />
              </h3>

              <InfoRow label="Content Locale" value={localeInfo.locale} icon="Globe" />
              <InfoRow label="UI Locale" value={localeInfo.uiLocale} icon="LocaleLanguage" />
              <InfoRow label="Text Direction" value={localeInfo.isRtl ? 'Right-to-Left (RTL)' : 'Left-to-Right (LTR)'} icon="TextAlignLeft" />

              {localeInfo.timeZone && (
                <>
                  <InfoRow label="Time Zone" value={localeInfo.timeZone.description} icon="Clock" />
                  <InfoRow label="UTC Offset" value={`${localeInfo.timeZone.offset} minutes`} icon="DateTime" />
                  <Label>
                    <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                    Current time: {new Intl.DateTimeFormat(localeInfo.locale, {
                      weekday: 'long',
                      year: 'numeric',
                      month: 'long',
                      day: 'numeric',
                      hour: '2-digit',
                      minute: '2-digit'
                    }).format(new Date())}
                  </Label>
                </>
              )}

              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Use with JavaScript Intl APIs for i18n (dates, numbers, currencies)
              </Label>
            </Stack>

            {/* Teams Context */}
            {hasTeamsContext && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="TeamsLogo" style={{ marginRight: '8px' }} />
                  Microsoft Teams Context
                  <Separator />
                </h3>
                <InfoRow label="Teams Supported" value="Yes" icon="CompletedSolid" />
                <InfoRow label="Teams Theme" value={teamsTheme ?? 'default'} icon="Color" />
              </Stack>
            )}

            {/* List Context */}
            {listInfo && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="BulletedList" style={{ marginRight: '8px' }} />
                  List Context
                  <Separator />
                </h3>
                <InfoRow label="List ID" value={listInfo.id} icon="Fingerprint" />
                <InfoRow label="List Title" value={listInfo.title} icon="FabricTextHighlight" />
                {listInfo.baseTemplate && (
                  <InfoRow label="Base Template" value={String(listInfo.baseTemplate)} icon="PageData" />
                )}
                {listInfo.isDocumentLibrary && (
                  <Label>
                    <Icon iconName="FabricFolder" style={{ marginRight: '4px', color: '#0078d4' }} />
                    Document Library
                  </Label>
                )}
              </Stack>
            )}

            {/* Hub Site */}
            {hubInfo?.isHubSite && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="NetworkTower" style={{ marginRight: '8px' }} />
                  Hub Site Association
                  <Separator />
                </h3>

                {hubInfo.error && (
                  <MessageBar messageBarType={MessageBarType.error}>
                    Failed to load hub site URL: {hubInfo.error.message}
                  </MessageBar>
                )}

                <InfoRow label="Hub Site ID" value={hubInfo.hubSiteId} icon="Fingerprint" />

                {hubInfo.isLoading ? (
                  <MessageBar messageBarType={MessageBarType.info}>
                    Loading hub site URL...
                  </MessageBar>
                ) : hubInfo.hubSiteUrl ? (
                  <InfoRow label="Hub Site URL" value={hubInfo.hubSiteUrl} icon="Link" />
                ) : null}
              </Stack>
            )}

          </Stack>
        </PivotItem>

        {/* TAB 3: Advanced */}
        <PivotItem headerText="Advanced" itemIcon="DeveloperTools">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: '16px' } }}>

            {/* Permissions Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Permissions" style={{ marginRight: '8px' }} />
                Permissions (Current Site)
                <Separator />
              </h3>
              <StatusBadge label="Can Manage Web" available={canManageWeb} />
              <StatusBadge label="Can Manage Lists" available={canManageLists} />
            </Stack>

            {/* Cross-Site Permissions Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Globe" style={{ marginRight: '8px' }} />
                Cross-Site Permissions
                <Separator />
              </h3>

              <Label>
                <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                Check permissions on a different site (auto-fetch when URL is provided)
              </Label>

              <TextField
                label="Target Site URL"
                value={crossSiteUrl || ''}
                onChange={(_, value) => setCrossSiteUrl(value?.trim() || undefined)}
                placeholder="https://contoso.sharepoint.com/sites/targetsite"
                description="Enter a site URL - permissions will be fetched automatically"
              />

              {crossSitePermissions.isLoading && (
                <MessageBar messageBarType={MessageBarType.info}>
                  Loading permissions from {crossSiteUrl}...
                </MessageBar>
              )}

              {crossSitePermissions.error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Error: {crossSitePermissions.error.message}
                </MessageBar>
              )}

              {crossSiteUrl && !crossSitePermissions.isLoading && !crossSitePermissions.error && (
                <Stack tokens={{ childrenGap: 1 }}>
                  <InfoRow
                    label="Target Site"
                    value={crossSiteUrl}
                    icon="Globe"
                  />
                  <StatusBadge
                    label="Can Manage Web (Cross-Site)"
                    available={crossSitePermissions.hasWebPermission(SPPermission.manageWeb)}
                  />
                  <StatusBadge
                    label="Can Manage Lists (Cross-Site)"
                    available={crossSitePermissions.hasWebPermission(SPPermission.manageLists)}
                  />
                  <StatusBadge
                    label="Can Add Items (Cross-Site)"
                    available={crossSitePermissions.hasWebPermission(SPPermission.addListItems)}
                  />
                </Stack>
              )}
            </Stack>

            {/* User Photo Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="ContactCard" style={{ marginRight: '8px' }} />
                User Photo Demo
                <Separator />
              </h3>

              {userPhoto.error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Load Error: {userPhoto.error.message}
                </MessageBar>
              )}

              {userPhoto.isLoading ? (
                <MessageBar messageBarType={MessageBarType.info}>Loading photo from Microsoft Graph...</MessageBar>
              ) : userPhoto.photoUrl ? (
                <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                  <img
                    src={userPhoto.photoUrl}
                    alt={displayName}
                    style={{
                      width: 120,
                      height: 120,
                      borderRadius: '50%',
                      objectFit: 'cover',
                      border: '3px solid #0078d4',
                      boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
                    }}
                  />
                  <Stack tokens={{ childrenGap: 1 }}>
                    <InfoRow label="Display Name" value={displayName} icon="Contact" />
                    <InfoRow label="Photo Size" value="240x240 (default)" icon="ResizeMouseMedium" />
                    <InfoRow label="Photo Format" value="Blob URL" icon="FileImage" />
                    <InfoRow label="Is Ready" value={userPhoto.isReady ? 'Yes' : 'No'} icon="CheckMark" />
                  </Stack>
                </Stack>
              ) : (
                <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                  <div style={{
                    width: 120,
                    height: 120,
                    borderRadius: '50%',
                    backgroundColor: '#0078d4',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    color: 'white',
                    fontSize: 48,
                    fontWeight: 'bold'
                  }}>
                    {displayName ? displayName.charAt(0).toUpperCase() : '?'}
                  </div>
                  <Label>
                    <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                    No profile photo available. Showing initials fallback.
                  </Label>
                </Stack>
              )}

              <PrimaryButton
                text={userPhoto.isLoading ? 'Loading...' : 'Reload Photo'}
                onClick={userPhoto.reload}
                disabled={userPhoto.isLoading}
                iconProps={{ iconName: 'Refresh' }}
              />

              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Photo loaded from Microsoft Graph API (/me/photos/240x240/$value)
              </Label>
              <Label>
                <Icon iconName="Warning" style={{ marginRight: '4px', color: '#d83b01' }} />
                Requires Microsoft Graph permissions: User.Read
              </Label>
              <Label>
                <Icon iconName="LightningBolt" style={{ marginRight: '4px', color: '#107c10' }} />
                Blob URL automatically cleaned up on unmount (memory safe)
              </Label>
            </Stack>

            {/* OneDrive AppData Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="OneDrive" style={{ marginRight: '8px' }} />
                OneDrive AppData Demo
                <Separator />
              </h3>

              {oneDriveData.error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Load Error: {oneDriveData.error.message}
                </MessageBar>
              )}

              {oneDriveData.writeError && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Write Error: {oneDriveData.writeError.message}
                </MessageBar>
              )}

              {!oneDriveData.isLoading && oneDriveData.isNotFound && !oneDriveData.data && !oneDriveData.error && (
                <MessageBar messageBarType={MessageBarType.info}>
                  File not found in OneDrive yet (isNotFound=true). Click &quot;Save to OneDrive&quot; to create it.
                </MessageBar>
              )}

              {oneDriveData.isLoading ? (
                <MessageBar messageBarType={MessageBarType.info}>Loading from OneDrive...</MessageBar>
              ) : oneDriveData.data ? (
                <Stack tokens={{ childrenGap: 1 }}>
                  <InfoRow label="Message" value={oneDriveData.data.message} icon="Message" />
                  <InfoRow label="Counter" value={String(oneDriveData.data.counter)} icon="NumberField" />
                  <InfoRow label="Last Updated" value={new Date(oneDriveData.data.timestamp).toLocaleString()} icon="DateTime" />
                  <InfoRow label="Is Ready" value={oneDriveData.isReady ? 'Yes' : 'No'} icon="CheckMark" />
                  <InfoRow label="Is Not Found" value={oneDriveData.isNotFound ? 'Yes' : 'No'} icon="BlockedSiteSolid12" />
                </Stack>
              ) : (
                <Label>
                  {oneDriveData.isNotFound
                    ? 'No file found yet. Click "Save to OneDrive" to create it.'
                    : 'No data loaded yet. Click "Load" to fetch from OneDrive.'}
                </Label>
              )}

              <TextField
                label="New Message"
                value={oneDriveMessage}
                onChange={(_, newValue) => setOneDriveMessage(newValue ?? '')}
                placeholder="Enter a message to save..."
                disabled={oneDriveData.isWriting || oneDriveData.isLoading}
              />

              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <PrimaryButton
                  text="Load from OneDrive"
                  onClick={handleLoadOneDrive}
                  disabled={oneDriveData.isLoading || oneDriveData.isWriting}
                  iconProps={{ iconName: 'CloudDownload' }}
                />
                <PrimaryButton
                  text={oneDriveData.isWriting ? 'Saving...' : 'Save to OneDrive'}
                  onClick={handleSaveOneDrive}
                  disabled={!oneDriveMessage || oneDriveData.isWriting || oneDriveData.isLoading}
                  iconProps={{ iconName: 'CloudUpload' }}
                />
              </Stack>

              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Data is stored in your OneDrive (appRoot:/useSPFxOneDriveAppData/test-data.json).
              </Label>
              <Label>
                <Icon iconName="Warning" style={{ marginRight: '4px', color: '#d83b01' }} />
                Requires Microsoft Graph permissions: Files.ReadWrite or Files.ReadWrite.AppFolder
              </Label>
            </Stack>

            {/* Tenant Property Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Org" style={{ marginRight: '8px' }} />
                Tenant Property Demo
                <Separator />
              </h3>

              {(tenantVersion.error || tenantCounter.error) && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Load Error: {tenantVersion.error?.message || tenantCounter.error?.message}
                </MessageBar>
              )}

              {(tenantVersion.writeError || tenantCounter.writeError) && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Write Error: {tenantVersion.writeError?.message || tenantCounter.writeError?.message}
                </MessageBar>
              )}

              {(tenantVersion.isLoading || tenantCounter.isLoading) ? (
                <MessageBar messageBarType={MessageBarType.info}>Loading from tenant app catalog...</MessageBar>
              ) : (tenantVersion.data || tenantCounter.data) ? (
                <Stack tokens={{ childrenGap: 1 }}>
                  <InfoRow label="Version" value={tenantVersion.data ?? '(not set)'} icon="ServerEnviroment" />
                  {tenantVersion.description && (
                    <Label>
                      <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                      {tenantVersion.description}
                    </Label>
                  )}
                  <InfoRow label="Counter" value={String(tenantCounter.data ?? 0)} icon="NumberField" />
                  {tenantCounter.description && (
                    <Label>
                      <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                      {tenantCounter.description}
                    </Label>
                  )}
                  <InfoRow label="Can Write" value={tenantVersion.canWrite ? 'Yes' : 'No'} icon="Permissions" />
                </Stack>
              ) : (
                <Label>No data loaded yet. Click &quot;Load&quot; to fetch from tenant properties.</Label>
              )}

              <TextField
                label="New Version"
                value={tenantVersionInput}
                onChange={(_, newValue) => setTenantVersionInput(newValue ?? '')}
                placeholder="e.g., 1.0.0"
                disabled={tenantVersion.isWriting || tenantVersion.isLoading || !tenantVersion.canWrite}
              />

              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <PrimaryButton
                  text="Load Properties"
                  onClick={handleLoadTenantProperty}
                  disabled={tenantVersion.isLoading || tenantCounter.isLoading}
                  iconProps={{ iconName: 'CloudDownload' }}
                />
                <PrimaryButton
                  text={tenantVersion.isWriting ? 'Saving...' : 'Save Version'}
                  onClick={handleSaveTenantVersion}
                  disabled={!tenantVersionInput || tenantVersion.isWriting || !tenantVersion.canWrite}
                  iconProps={{ iconName: 'Save' }}
                />
                <PrimaryButton
                  text={tenantCounter.isWriting ? 'Incrementing...' : 'Increment Counter'}
                  onClick={handleIncrementTenantCounter}
                  disabled={tenantCounter.isWriting || !tenantCounter.canWrite}
                  iconProps={{ iconName: 'Add' }}
                />
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <DefaultButton
                  text="Remove Version"
                  onClick={() => handleRemoveTenantProperty('version')}
                  disabled={tenantVersion.isWriting || !tenantVersion.canWrite || !tenantVersion.data}
                  iconProps={{ iconName: 'Delete' }}
                />
                <DefaultButton
                  text="Remove Counter"
                  onClick={() => handleRemoveTenantProperty('counter')}
                  disabled={tenantCounter.isWriting || !tenantCounter.canWrite || !tenantCounter.data}
                  iconProps={{ iconName: 'Delete' }}
                />
              </Stack>

              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Properties are stored tenant-wide in SharePoint StorageEntity. All users can read, only admins can write.
              </Label>
              <Label>
                <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                Keys: spfx-toolkit-test-version (string), spfx-toolkit-test-counter (number)
              </Label>
              <Label>
                <Icon iconName="Warning" style={{ marginRight: '4px', color: '#d83b01' }} />
                Requires: Tenant app catalog provisioned, Manage Web permissions to write
              </Label>

              {!tenantVersion.canWrite && (
                <MessageBar messageBarType={MessageBarType.info}>
                  ℹ️ You don&apos;t have permission to modify tenant properties. Contact your SharePoint administrator.
                </MessageBar>
              )}
            </Stack>


            {/* Storage Demo Cards */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Database" style={{ marginRight: '8px' }} />
                Session Storage Demo
                <Separator />
              </h3>

              <InfoRow label="Current Value" value={sessionStorage.value || '(empty)'} icon="Variable" />
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <TextField
                  value={sessionStorageInput}
                  onChange={(_, newValue) => setSessionStorageInput(newValue ?? '')}
                  placeholder="Enter value..."
                  styles={{ root: { flexGrow: 1 } }}
                />
                <PrimaryButton text="Save" onClick={handleSaveSessionStorage} iconProps={{ iconName: 'Save' }} />
                <DefaultButton text="Clear" onClick={() => sessionStorage.remove()} iconProps={{ iconName: 'Delete' }} />
              </Stack>
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Persists only for current tab/session
              </Label>
            </Stack>

            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Save" style={{ marginRight: '8px' }} />
                Local Storage Demo
                <Separator />
              </h3>
              <InfoRow label="Current Value" value={localStorage.value || '(empty)'} icon="Variable" />
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <TextField
                  value={localStorageInput}
                  onChange={(_, newValue) => setLocalStorageInput(newValue ?? '')}
                  placeholder="Enter value..."
                  styles={{ root: { flexGrow: 1 } }}
                />
                <PrimaryButton text="Save" onClick={handleSaveLocalStorage} iconProps={{ iconName: 'Save' }} />
                <DefaultButton text="Clear" onClick={() => localStorage.remove()} iconProps={{ iconName: 'Delete' }} />
              </Stack>
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Persists across sessions and page reloads
              </Label>
            </Stack>

            {/* Performance Test Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="SpeedHigh" style={{ marginRight: '8px' }} />
                Performance Test
                <Separator />
              </h3>
              {performanceResult && (
                <MessageBar messageBarType={MessageBarType.info}>
                  {performanceResult}
                </MessageBar>
              )}
              <PrimaryButton
                text="Run Performance Test"
                onClick={handlePerformanceTest}
                iconProps={{ iconName: 'LightningBolt' }}
              />
            </Stack>

            {/* Logger Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Code" style={{ marginRight: '8px' }} />
                Logger Demo
                <Separator />
              </h3>
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <DefaultButton
                  text="Info"
                  onClick={() => handleLog('info', 'This is an info message')}
                  iconProps={{ iconName: 'Info' }}
                />
                <DefaultButton
                  text="Warning"
                  onClick={() => handleLog('warning', 'This is a warning message')}
                  iconProps={{ iconName: 'Warning' }}
                />
                <DefaultButton
                  text="Error"
                  onClick={() => handleLog('error', 'This is an error message')}
                  iconProps={{ iconName: 'Error' }}
                />
              </Stack>
              {logMessages.length > 0 && (
                <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginTop: '8px' } }}>
                  <Label>Recent Logs (check browser console):</Label>
                  {logMessages.map((log, idx) => {
                    const levelClass = log.level === 'info' ? styles.info :
                      log.level === 'warning' ? styles.warning :
                        log.level === 'error' ? styles.error :
                          styles.verbose;
                    return (
                      <div key={idx} className={`${styles.logMessage} ${levelClass}`}>
                        <strong>[{log.level.toUpperCase()}]</strong>: {log.message}
                      </div>
                    );
                  })}
                </Stack>
              )}
            </Stack>

            {/* HTTP Clients Status Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="CloudUpload" style={{ marginRight: '8px' }} />
                HTTP Clients Status
                <Separator />
              </h3>
              <StatusBadge label="SPHttpClient" available={!!spHttpClient} />
              <StatusBadge label="MSGraphClient" available={!!msGraphClient} />
              <StatusBadge label="AadHttpClient" available={!!aadHttpClient} />
            </Stack>

            {/* Theme Information Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Color" style={{ marginRight: '8px' }} />
                Theme Information
                <Separator />
              </h3>
              <InfoRow label="Is Dark Theme" value={isDarkTheme ? 'Yes' : 'No'} icon="Brightness" />
              <InfoRow label="Body Background" value={theme?.semanticColors?.bodyBackground ?? 'N/A'} icon="FabricFolderFill" />
              <InfoRow label="Body Text" value={theme?.semanticColors?.bodyText ?? 'N/A'} icon="Font" />
              <InfoRow label="Link Color" value={theme?.semanticColors?.link ?? 'N/A'} icon="Link" />
            </Stack>

            {/* Fluent UI 9 Theme Information Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Color" style={{ marginRight: '8px' }} />
                Fluent UI 9 Theme Information
                <Separator />
              </h3>
              <InfoRow label="Is Teams Context" value={fluent9ThemeInfo.isTeams ? 'Yes' : 'No'} icon="TeamsLogo" />
              {fluent9ThemeInfo.isTeams && fluent9ThemeInfo.teamsTheme && (
                <InfoRow label="Teams Theme" value={fluent9ThemeInfo.teamsTheme} icon="Color" />
              )}
              <details className={styles.detailsSection}>
                <summary>
                  Click to expand Fluent UI 9 Theme (JSON)
                </summary>
                <pre>
                  {safeStringify(fluent9ThemeInfo.theme, 2)}
                </pre>
              </details>
            </Stack>

            {/* Advanced Diagnostics Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="DeveloperTools" style={{ marginRight: '8px' }} />
                Advanced Diagnostics
                <Separator />
              </h3>
              <InfoRow label="Correlation ID" value={correlationInfo.correlationId} icon="TrackersMirrored" />
              <InfoRow label="Tenant ID" value={correlationInfo.tenantId} icon="CityNext" />
              <InfoRow label="ServiceScope" value={serviceScope ? 'Available' : 'N/A'} icon="Settings" />
            </Stack>

            {/* Page Context Raw Data */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="FileCode" style={{ marginRight: '8px' }} />
                Page Context (Raw JSON)
                <Separator />
              </h3>
              <details className={styles.detailsSection}>
                <summary>
                  Click to expand/collapse
                </summary>
                <pre>
                  {safeStringify(pageContext, 2)}
                </pre>
              </details>
            </Stack>

          </Stack>
        </PivotItem>

        {/* TAB 4: HttpClient Example */}
        <PivotItem headerText="HttpClient" itemIcon="Cloud">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: '16px' } }}>

            <MessageBar messageBarType={MessageBarType.info}>
              <strong>useSPFxHttpClient Hook Example</strong>
              <br />
              This hook provides access to the generic HttpClient for calling external APIs (non-SharePoint).
              For SharePoint REST API calls, use <strong>useSPFxSPHttpClient</strong> instead.
            </MessageBar>

            <HttpClientDemo />

          </Stack>
        </PivotItem>

        {/* TAB 5: PnPjs Examples */}
        <PivotItem headerText="PnPjs" itemIcon="CloudDownload">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: '16px' } }}>

            {/* Example 1: useSPFxPnPContext */}
            <PnPContextDemo />

            {/* Example 2: useSPFxPnP */}
            <PnPOperationsDemo />

            {/* Example 3: useSPFxPnPList */}
            <PnPListDemo />

          </Stack>
        </PivotItem>

        {/* TAB 6: Search Examples */}
        <PivotItem headerText="Search" itemIcon="Search">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: '16px' } }}>

            <MessageBar messageBarType={MessageBarType.info}>
              <strong>useSPFxPnPSearch Hook Examples</strong>
              <br />
              These examples demonstrate SharePoint Search capabilities using the native PnPjs SearchQueryBuilder API.
            </MessageBar>

            {/* Example 1: Basic Search */}
            <PnPSearchBasicDemo />

            {/* Example 2: Advanced Search with Builder */}
            <PnPSearchAdvancedDemo />

            {/* Example 3: Faceted Search with Refiners */}
            <PnPSearchRefinersDemo />

            {/* Example 4: Search Suggestions (Autocomplete) */}
            <PnPSearchSuggestionsDemo />

          </Stack>
        </PivotItem>

      </Pivot>

    </section>
  );
};

export default SpFxReactToolkitTest;
