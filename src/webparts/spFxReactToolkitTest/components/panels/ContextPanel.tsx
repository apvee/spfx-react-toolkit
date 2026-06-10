import * as React from 'react';
import {
  MessageBar,
  MessageBarType,
  Stack,
  TextField,
} from '@fluentui/react';
import { SPPermission } from '@microsoft/sp-page-context';
import {
  useSPFxContainerInfo,
  useSPFxContainerSize,
  useSPFxContext,
  useSPFxCorrelationInfo,
  useSPFxCrossSitePermissions,
  useSPFxDisplayMode,
  useSPFxEnvironmentInfo,
  useSPFxFluent9ThemeInfo,
  useSPFxHubSiteInfo,
  useSPFxInstanceInfo,
  useSPFxIsEdit,
  useSPFxListInfo,
  useSPFxLocaleInfo,
  useSPFxPageContext,
  useSPFxPageType,
  useSPFxPermissions,
  useSPFxServiceScope,
  useSPFxSiteInfo,
  useSPFxTeams,
  useSPFxThemeInfo,
  useSPFxUserInfo,
} from '../../../../hooks';
import {
  DemoCard,
  InfoGrid,
  JsonDetails,
  StatusBadge,
} from '../shared';

const bool = (value: boolean): string => value ? 'Yes' : 'No';

const ContextPanel: React.FC = () => {
  const contextValue = useSPFxContext();
  const pageContext = useSPFxPageContext();
  const instance = useSPFxInstanceInfo();
  const displayMode = useSPFxDisplayMode();
  const isEdit = useSPFxIsEdit();
  const environment = useSPFxEnvironmentInfo();
  const user = useSPFxUserInfo();
  const site = useSPFxSiteInfo();
  const list = useSPFxListInfo();
  const locale = useSPFxLocaleInfo();
  const hub = useSPFxHubSiteInfo();
  const teams = useSPFxTeams();
  const pageType = useSPFxPageType();
  const correlation = useSPFxCorrelationInfo();
  const permissions = useSPFxPermissions();
  const containerInfo = useSPFxContainerInfo();
  const containerSize = useSPFxContainerSize();
  const theme = useSPFxThemeInfo();
  const fluent9 = useSPFxFluent9ThemeInfo();
  const serviceScope = useSPFxServiceScope();
  const [crossSiteUrl, setCrossSiteUrl] = React.useState<string | undefined>();
  const crossSite = useSPFxCrossSitePermissions(crossSiteUrl);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <DemoCard title="Instance and Host" iconName="WebAppBuilderFragment">
        <InfoGrid
          rows={[
            { label: 'Context Kind', value: contextValue.kind, icon: 'CubeShape' },
            { label: 'Instance ID', value: instance.id, icon: 'Fingerprint' },
            { label: 'Component Kind', value: instance.kind, icon: 'CubeShape' },
            { label: 'Display Mode', value: displayMode.isEdit ? 'Edit' : 'Read', icon: 'Edit' },
            { label: 'useSPFxIsEdit', value: bool(isEdit), icon: 'EditContact' },
            { label: 'Environment', value: environment.type, icon: 'Globe' },
            { label: 'Page Type', value: pageType.pageType, icon: 'Page' },
            { label: 'ServiceScope', value: serviceScope.serviceScope ? 'Available' : 'Unavailable', icon: 'Settings' },
          ]}
        />
      </DemoCard>

      <DemoCard title="User, Site, and Locale" iconName="People">
        <InfoGrid
          rows={[
            { label: 'User Name', value: user.displayName, icon: 'Contact' },
            { label: 'User Email', value: user.email, icon: 'Mail' },
            { label: 'External User', value: bool(user.isExternal), icon: 'ContactCard' },
            { label: 'Web Title', value: site.title, icon: 'CityNext' },
            { label: 'Web URL', value: site.webUrl, icon: 'Link' },
            { label: 'Site URL', value: site.siteUrl, icon: 'Home' },
            { label: 'Classification', value: site.siteClassification, icon: 'Tag' },
            { label: 'Locale', value: locale.locale, icon: 'LocaleLanguage' },
            { label: 'UI Locale', value: locale.uiLocale, icon: 'Globe' },
            { label: 'RTL', value: bool(locale.isRtl), icon: 'TextAlignLeft' },
            { label: 'Time Zone', value: locale.timeZone?.description, icon: 'Clock' },
          ]}
        />
      </DemoCard>

      <DemoCard title="List, Hub, Teams, and Theme" iconName="FabricFolder">
        <InfoGrid
          rows={[
            { label: 'List ID', value: list?.id, icon: 'Fingerprint' },
            { label: 'List Title', value: list?.title, icon: 'BulletedList' },
            { label: 'Document Library', value: list ? bool(list.isDocumentLibrary ?? false) : undefined, icon: 'FabricFolder' },
            { label: 'Hub Site', value: bool(hub.isHubSite), icon: 'NetworkTower' },
            { label: 'Hub Site ID', value: hub.hubSiteId, icon: 'Fingerprint' },
            { label: 'Hub Site URL', value: hub.hubSiteUrl, icon: 'Link' },
            { label: 'Teams Supported', value: bool(teams.supported), icon: 'TeamsLogo' },
            { label: 'Teams Theme', value: teams.theme, icon: 'Color' },
            { label: 'SPFx Theme Inverted', value: theme ? bool(theme.isInverted ?? false) : undefined, icon: 'Brightness' },
            { label: 'Fluent 9 Teams Theme', value: fluent9.teamsTheme, icon: 'Color' },
          ]}
        />
        {hub.error && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Hub site lookup failed: {hub.error.message}
          </MessageBar>
        )}
      </DemoCard>

      <DemoCard title="Permissions and Correlation" iconName="Permissions">
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <StatusBadge label="Manage Web" available={permissions.hasWebPermission(SPPermission.manageWeb)} />
          <StatusBadge label="Manage Lists" available={permissions.hasWebPermission(SPPermission.manageLists)} />
          <StatusBadge label="Add List Items" available={permissions.hasWebPermission(SPPermission.addListItems)} />
        </Stack>
        <InfoGrid
          rows={[
            { label: 'Correlation ID', value: correlation.correlationId, icon: 'TrackersMirrored' },
            { label: 'Tenant ID', value: correlation.tenantId, icon: 'CityNext' },
            { label: 'Site Permissions', value: permissions.sitePermissions ? 'Available' : 'Unavailable', icon: 'Permissions' },
            { label: 'Web Permissions', value: permissions.webPermissions ? 'Available' : 'Unavailable', icon: 'Permissions' },
            { label: 'List Permissions', value: permissions.listPermissions ? 'Available' : 'Unavailable', icon: 'Permissions' },
          ]}
        />
      </DemoCard>

      <DemoCard title="Cross-Site Permissions" iconName="Globe">
        <TextField
          label="Target site URL"
          value={crossSiteUrl ?? ''}
          onChange={(_, value) => setCrossSiteUrl(value?.trim() || undefined)}
          placeholder="https://contoso.sharepoint.com/sites/target"
          description="Leave empty to keep this hook idle."
        />
        {crossSite.isLoading && (
          <MessageBar messageBarType={MessageBarType.info}>
            Loading permissions...
          </MessageBar>
        )}
        {crossSite.error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {crossSite.error.message}
          </MessageBar>
        )}
        {crossSiteUrl && !crossSite.isLoading && !crossSite.error && (
          <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
            <StatusBadge label="Cross-site Manage Web" available={crossSite.hasWebPermission(SPPermission.manageWeb)} />
            <StatusBadge label="Cross-site Manage Lists" available={crossSite.hasWebPermission(SPPermission.manageLists)} />
            <StatusBadge label="Cross-site Add Items" available={crossSite.hasWebPermission(SPPermission.addListItems)} />
          </Stack>
        )}
      </DemoCard>

      <DemoCard title="Container" iconName="Resize">
        <InfoGrid
          rows={[
            { label: 'Container Width', value: `${containerSize.width}px`, icon: 'ArrowRight' },
            { label: 'Container Height', value: `${containerSize.height}px`, icon: 'ArrowUp' },
            { label: 'Container Size', value: containerSize.size, icon: 'FitPage' },
            { label: 'Container Element', value: containerInfo.element ? 'Available' : 'Unavailable', icon: 'DOM' },
            { label: 'Tracked Size', value: containerInfo.size ? 'Available' : 'Unavailable', icon: 'RadioBullet' },
          ]}
        />
      </DemoCard>

      <JsonDetails label="Page context JSON" value={pageContext} />
      <JsonDetails label="Fluent UI 9 theme JSON" value={fluent9.theme} />
    </Stack>
  );
};

export default ContextPanel;
