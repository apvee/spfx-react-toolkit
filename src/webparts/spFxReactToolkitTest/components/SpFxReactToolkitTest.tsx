import * as React from 'react';
import styles from './SpFxReactToolkitTest.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
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

// Import all hooks from SPFx React Toolkit
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
} from '../../../hooks';
import { SPPermission } from '@microsoft/sp-page-context';

interface IWebPartProps {
  description: string;
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

// Helper component for displaying info rows
const InfoRow: React.FC<{ label: string; value: string | undefined; icon?: string }> = ({ label, value, icon }) => (
  <div className={styles.infoRow}>
    {icon && <Icon iconName={icon} />}
    <Label>{label}:</Label>
    <span>{value || 'N/A'}</span>
  </div>
);

// Helper component for status badges
const StatusBadge: React.FC<{ available: boolean; label: string }> = ({ available, label }) => (
  <div className={`${styles.statusBadge} ${available ? styles.success : styles.error}`}>
    <Icon iconName={available ? 'Completed' : 'ErrorBadge'} />
    {label}
  </div>
);

const SpFxReactToolkitTest: React.FC = () => {
  // Core Properties & Display
  const { properties, setProperties } = useSPFxProperties<IWebPartProps>();
  const { isEdit } = useSPFxDisplayMode();
  const { id, kind } = useSPFxInstanceInfo();

  // Theme & Environment
  const theme = useSPFxThemeInfo();
  const { type: envType, isLocal } = useSPFxEnvironmentInfo();
  const isDarkTheme = theme?.isInverted ?? false;

  // User & Site Info
  const { displayName } = useSPFxUserInfo();
  const { title: siteTitle, webUrl, siteClassification } = useSPFxSiteInfo();
  const { uiLocale } = useSPFxLocaleInfo();

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

  // Local state for interactive demos
  const [descriptionInput, setDescriptionInput] = React.useState(properties?.description ?? '');
  const [sessionStorageInput, setSessionStorageInput] = React.useState('');
  const [localStorageInput, setLocalStorageInput] = React.useState('');
  const [showMessage, setShowMessage] = React.useState(false);
  const [messageText, setMessageText] = React.useState('');
  const [performanceResult, setPerformanceResult] = React.useState<string>('');
  const [logMessages, setLogMessages] = React.useState<Array<{ level: string; message: string }>>([]);

  // Sync descriptionInput when properties change
  React.useEffect(() => {
    setDescriptionInput(properties?.description ?? '');
  }, [properties?.description]);

  // Environment message (calculated)
  const environmentMessage = React.useMemo(() => {
    if (isLocal) {
      return hasTeamsContext ? 'Local Workbench (Teams)' : 'Local Workbench (SharePoint)';
    }
    if (hasTeamsContext) {
      return 'Microsoft Teams';
    }
    return 'SharePoint Online';
  }, [isLocal, hasTeamsContext]);

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

  return (
    <section className={`${styles.spFxReactToolkitTest} ${hasTeamsContext ? styles.teams : ''}`}>
      {/* Welcome Section */}
      <div className={styles.welcome}>
        <img
          alt=""
          src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')}
          className={styles.welcomeImage}
        />
        <h2>Well done, {escape(displayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(properties?.description ?? 'No description')}</strong></div>
      </div>

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
              <InfoRow label="UI Locale" value={uiLocale} icon="LocaleLanguage" />
              {siteClassification && (
                <InfoRow label="Classification" value={siteClassification} icon="Tag" />
              )}
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
                <InfoRow label="Hub Site ID" value={hubInfo.hubSiteId} icon="Fingerprint" />
                <InfoRow label="Hub Site URL" value={hubInfo.hubSiteUrl} icon="Link" />
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
                Permissions
                <Separator />
              </h3>
              <StatusBadge label="Can Manage Web" available={canManageWeb} />
              <StatusBadge label="Can Manage Lists" available={canManageLists} />
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

      </Pivot>

    </section>
  );
};

export default SpFxReactToolkitTest;
