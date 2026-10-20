/** Regression snapshot captured before single-subject consolidation; not product UI. */
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';
import { Field, Input, MessageBar, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { RelativeTime, useSPFxContainerInfo, useSPFxContainerSize, useSPFxContext, useSPFxCorrelationInfo, useSPFxCrossSitePermissions, useSPFxDisplayMode, useSPFxEnvironmentInfo, useSPFxFluent9ThemeInfo, useSPFxHubSiteInfo, useSPFxInstanceInfo, useSPFxIsEdit, useSPFxListInfo, useSPFxLocaleInfo, useSPFxPageContext, useSPFxPageType, useSPFxPermissions, useSPFxServiceScope, useSPFxSiteInfo, useSPFxTeams, useSPFxThemeInfo, useSPFxUserInfo, } from '@apvee/spfx-react-toolkit';
import { DemoCard, InfoGrid, JsonDetails, StatusBadge, } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';
const bool = (value: boolean): string => value ? 'Yes' : 'No';
const ContextPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
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
  return (<div className={demoStyles.root}>
    <DemoCard contentPlacement="configuration" title="Instance and Host" iconName="WebAppBuilderFragment">
      <InfoGrid rows={[
        { label: 'Context Kind', value: contextValue.kind, icon: 'CubeShape' },
        { label: 'Instance ID', value: instance.id, icon: 'Fingerprint' },
        { label: 'Component Kind', value: instance.kind, icon: 'CubeShape' },
        { label: 'Display Mode', value: displayMode.isEdit ? 'Edit' : 'Read', icon: 'Edit' },
        { label: 'useSPFxIsEdit', value: bool(isEdit), icon: 'EditContact' },
        { label: 'Environment', value: environment.type, icon: 'Globe' },
        { label: 'Page Type', value: pageType.pageType, icon: 'Page' },
        { label: 'ServiceScope', value: serviceScope.serviceScope ? 'Available' : 'Unavailable', icon: 'Settings' },
      ]} />
    </DemoCard>

    <DemoCard contentPlacement="configuration" title="User, Site, and Locale" iconName="People">
      <InfoGrid rows={[
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
        { label: 'Representative IANA Zone', value: locale.ianaTimeZone ?? 'Unavailable', icon: 'Clock' },
      ]} />
    </DemoCard>

    <DemoCard contentPlacement="configuration" title="Site Time Zone — Fixed Winter/Summer Instants" iconName="Clock">
      {locale.ianaTimeZone ? <>
        <p>Representative standard zone: {locale.ianaTimeZone}. Focus or hover each time to compare its full date.</p>
        <p>Winter (12:00 UTC): <RelativeTime data-relative-time-context="winter" dateTime="2026-01-15T12:00:00Z" referenceTime="2026-01-15T13:00:00Z" locale={locale.uiLocale} timeZone={locale.ianaTimeZone} />
        </p>
        <p>Summer (12:00 UTC): <RelativeTime data-relative-time-context="summer" dateTime="2026-07-15T12:00:00Z" referenceTime="2026-07-15T13:00:00Z" locale={locale.uiLocale} timeZone={locale.ianaTimeZone} />
        </p>
        <p>This mapping represents a standard ID group; it does not identify a user's city or custom TIMEZONE.XML rules.</p>
      </> : <p>IANA mapping unavailable. No site-zone full date is shown; an omitted timeZone would use the runtime zone.</p>}
    </DemoCard>

    <DemoCard contentPlacement="configuration" title="List, Hub, Teams, and Theme" iconName="FabricFolder">
      <InfoGrid rows={[
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
      ]} />
      {hub.error && (<MessageBar intent="warning">
        <MessageBarBody>
          Hub site lookup failed: {hub.error.message}
        </MessageBarBody>
      </MessageBar>)}
    </DemoCard>

    <DemoCard contentPlacement="configuration" title="Permissions and Correlation" iconName="Permissions">
      <div className={demoStyles.actions}>
        <StatusBadge label="Manage Web" available={permissions.hasWebPermission(SPPermission.manageWeb)} />
        <StatusBadge label="Manage Lists" available={permissions.hasWebPermission(SPPermission.manageLists)} />
        <StatusBadge label="Add List Items" available={permissions.hasWebPermission(SPPermission.addListItems)} />
      </div>
      <InfoGrid rows={[
        { label: 'Correlation ID', value: correlation.correlationId, icon: 'TrackersMirrored' },
        { label: 'Tenant ID', value: correlation.tenantId, icon: 'CityNext' },
        { label: 'Site Permissions', value: permissions.sitePermissions ? 'Available' : 'Unavailable', icon: 'Permissions' },
        { label: 'Web Permissions', value: permissions.webPermissions ? 'Available' : 'Unavailable', icon: 'Permissions' },
        { label: 'List Permissions', value: permissions.listPermissions ? 'Available' : 'Unavailable', icon: 'Permissions' },
      ]} />
    </DemoCard>

    <DemoCard contentPlacement="configuration" title="Cross-Site Permissions" iconName="Globe" controls={<>
      <Field label={"Target site URL"} hint={"Leave empty to keep this hook idle."}>
        <Input value={crossSiteUrl ?? ''} onChange={(_, data) => setCrossSiteUrl(data.value?.trim() || undefined)} placeholder="https://contoso.sharepoint.com/sites/target" />
      </Field>
    </>}>
      {crossSite.isLoading && (<MessageBar intent="info">
        <MessageBarBody>
          Loading permissions...
        </MessageBarBody>
      </MessageBar>)}
      {crossSite.error && (<MessageBar intent="error">
        <MessageBarBody>
          {crossSite.error.message}
        </MessageBarBody>
      </MessageBar>)}
      {crossSiteUrl && !crossSite.isLoading && !crossSite.error && (<div className={demoStyles.actions}>
        <StatusBadge label="Cross-site Manage Web" available={crossSite.hasWebPermission(SPPermission.manageWeb)} />
        <StatusBadge label="Cross-site Manage Lists" available={crossSite.hasWebPermission(SPPermission.manageLists)} />
        <StatusBadge label="Cross-site Add Items" available={crossSite.hasWebPermission(SPPermission.addListItems)} />
      </div>)}
    </DemoCard>

    <DemoCard contentPlacement="configuration" title="Container" iconName="Resize" diagnostics={<>
      <JsonDetails label="Page context JSON" value={pageContext} />
      <JsonDetails label="Fluent UI 9 theme JSON" value={fluent9.theme} />
    </>}>
      <InfoGrid rows={[
        { label: 'Container Width', value: `${containerSize.width}px`, icon: 'ArrowRight' },
        { label: 'Container Height', value: `${containerSize.height}px`, icon: 'ArrowUp' },
        { label: 'Container Size', value: containerSize.size, icon: 'FitPage' },
        { label: 'Container Element', value: containerInfo.element ? 'Available' : 'Unavailable', icon: 'DOM' },
        { label: 'Tracked Size', value: containerInfo.size ? 'Available' : 'Unavailable', icon: 'RadioBullet' },
      ]} />
    </DemoCard>

  </div>);
};
export default ContextPanel;
