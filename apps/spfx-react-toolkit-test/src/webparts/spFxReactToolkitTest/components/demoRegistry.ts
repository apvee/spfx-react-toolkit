import * as React from 'react';

import type { DemoCoverageItem, DemoRegistryEntry } from './demoRegistry.types';
export type { DemoCoverageKind, DemoCoverageItem, DemoRegistryEntry } from './demoRegistry.types';

const lazyPanel = (
  loader: () => Promise<{ default: React.ComponentType }>
): React.LazyExoticComponent<React.ComponentType> => React.lazy(loader);

/** Temporary panels only: these entries are not generated features or executed test results. */
export const demoRegistry: readonly DemoRegistryEntry[] = [
  {
    key: 'file-folder-picker', title: 'File and folder picker', iconName: 'Folder',
    description: 'Microsoft v8 hosted picker with verified SharePoint and OneDrive Business boundaries.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-file-folder-picker' */ './panels/FileFolderPickerPanel')),
    coverage: [
      { symbol: 'FileFolderPicker', kind: 'read', notes: 'One live hosted v8 picker; custom browse dictionary, source/kind/category/cardinality/presentation, optional colorful tag avatars and consumer-owned JSON references.' },
      { symbol: 'PropertyPaneFileFolderPicker', kind: 'host', notes: 'Native direct JSON arrays, independent hosted presentation, optional colorful tag avatars and category/selection policy; source cascade resets explicitly.' },
    ],
  },
  {
    key: 'pickers', title: 'SharePoint pickers', iconName: 'Search',
    description: 'Seven live React pickers and native Property Pane JSON fields.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-pickers' */ './panels/PickersPanel')),
    coverage: [
      { symbol: 'SitePicker', kind: 'read', notes: 'Live authenticated picker; native/default or per-instance search/status dictionary, configurable presentation, optional colorful tag avatars, selection, consumer models and JSON diagnostics.' },
      { symbol: 'ListPicker', kind: 'read', notes: 'Live authenticated picker; native/default or per-instance search/status dictionary, configurable presentation, optional colorful tag avatars, selection, consumer models and JSON diagnostics.' },
      { symbol: 'FieldPicker', kind: 'read', notes: 'Live authenticated picker; native/default or per-instance search/status dictionary, configurable presentation, optional colorful tag avatars, selection, consumer models and JSON diagnostics.' },
      { symbol: 'ContentTypePicker', kind: 'read', notes: 'Live site collection/list metadata picker with per-instance strings; complete catalog, local filters, optional colorful tag avatars and JSON diagnostics.' },
      { symbol: 'ViewPicker', kind: 'read', notes: 'Live authenticated picker; native/default or per-instance search/status dictionary, configurable presentation, optional colorful tag avatars, selection, consumer models and JSON diagnostics.' },
      { symbol: 'ListItemPicker', kind: 'read', notes: 'Live authenticated picker; native/default or per-instance search/status dictionary, configurable presentation, optional colorful tag avatars, selection, consumer models and JSON diagnostics.' },
      { symbol: 'PrincipalPicker', kind: 'read', notes: 'Live authenticated picker; native/default or per-instance search/status dictionary, configurable presentation, optional colorful tag avatars, selection, consumer models and JSON diagnostics.' },
      { symbol: 'PropertyPaneSitePicker', kind: 'host', notes: 'Native custom field in getPropertyPaneConfiguration; stable key and JSON value arrays.' },
      { symbol: 'PropertyPaneListPicker', kind: 'host', notes: 'Native custom field in getPropertyPaneConfiguration; stable key and JSON value arrays.' },
      { symbol: 'PropertyPaneFieldPicker', kind: 'host', notes: 'Native custom field in getPropertyPaneConfiguration; stable key and JSON value arrays.' },
      { symbol: 'PropertyPaneContentTypePicker', kind: 'host', notes: 'Native site/list content type reference arrays with stable key and source cascade revision.' },
      { symbol: 'PropertyPaneViewPicker', kind: 'host', notes: 'Native custom field in getPropertyPaneConfiguration; stable key and JSON value arrays.' },
      { symbol: 'PropertyPaneListItemPicker', kind: 'host', notes: 'Native custom field in getPropertyPaneConfiguration; stable key and JSON value arrays.' },
      { symbol: 'PropertyPanePrincipalPicker', kind: 'host', notes: 'Native custom field in getPropertyPaneConfiguration; stable key and JSON value arrays.' },
    ],
  },

  {
    key: 'providers',
    title: 'Providers',
    iconName: 'PlugConnected',
    description: 'Host-specific provider integration points exposed by the package.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-providers' */ './panels/ProvidersPanel')),
    coverage: [
      { symbol: 'SPFxWebPartProvider', kind: 'host', notes: 'Active host for this webpart.' },
      { symbol: 'SPFxApplicationCustomizerProvider', kind: 'host', notes: 'Shown by the sample extension.' },
      { symbol: 'SPFxFieldCustomizerProvider', kind: 'host', notes: 'Host-specific provider export verified; requires Field Customizer host for mounting.' },
      { symbol: 'SPFxListViewCommandSetProvider', kind: 'host', notes: 'Host-specific provider export verified; requires Command Set host for mounting.' },
    ],
  },
  {
    key: 'context',
    title: 'Context',
    iconName: 'PageHeader',
    description: 'SPFx context, page, user, site, permissions, Teams, theme, and container hooks.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-context' */ './panels/ContextPanel')),
    coverage: [
      { symbol: 'useSPFxContext', kind: 'snapshot', notes: 'Provider context value.' },
      { symbol: 'useSPFxPageContext', kind: 'snapshot', notes: 'Native SPFx page context.' },
      { symbol: 'useSPFxInstanceInfo', kind: 'snapshot', notes: 'Instance id and host kind.' },
      { symbol: 'useSPFxDisplayMode', kind: 'snapshot', notes: 'Read/edit display mode.' },
      { symbol: 'useSPFxIsEdit', kind: 'snapshot', notes: 'Edit-mode shortcut.' },
      { symbol: 'useSPFxEnvironmentInfo', kind: 'snapshot', notes: 'SharePoint, Teams, local environment.' },
      { symbol: 'useSPFxUserInfo', kind: 'snapshot', notes: 'Current user metadata.' },
      { symbol: 'useSPFxSiteInfo', kind: 'snapshot', notes: 'Current site and web metadata.' },
      { symbol: 'useSPFxListInfo', kind: 'snapshot', notes: 'List/library context when available.' },
      { symbol: 'useSPFxLocaleInfo', kind: 'snapshot', notes: 'Locale, UI locale, and time zone.' },
      { symbol: 'useSPFxHubSiteInfo', kind: 'read', notes: 'Hub association and URL lookup.' },
      { symbol: 'useSPFxTeams', kind: 'snapshot', notes: 'Teams context and theme state.' },
      { symbol: 'useSPFxPageType', kind: 'snapshot', notes: 'Modern page/list/form classification.' },
      { symbol: 'useSPFxCorrelationInfo', kind: 'snapshot', notes: 'Correlation and tenant identifiers.' },
      { symbol: 'useSPFxPermissions', kind: 'snapshot', notes: 'Current site permission helpers.' },
      { symbol: 'useSPFxCrossSitePermissions', kind: 'read', notes: 'Optional target site permission check.' },
      { symbol: 'useSPFxContainerInfo', kind: 'snapshot', notes: 'Container element and tracked size.' },
      { symbol: 'useSPFxContainerSize', kind: 'snapshot', notes: 'Responsive size bucket.' },
      { symbol: 'useSPFxThemeInfo', kind: 'snapshot', notes: 'SPFx theme.' },
      { symbol: 'useSPFxFluent9ThemeInfo', kind: 'snapshot', notes: 'Fluent UI 9 theme adapter.' },
      { symbol: 'useSPFxServiceScope', kind: 'snapshot', notes: 'SPFx service scope access.' },
    ],
  },
  {
    key: 'runtime',
    title: 'Runtime',
    iconName: 'DeveloperTools',
    description: 'Webpart properties, scoped browser storage, logging, and performance helpers.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-runtime' */ './panels/RuntimePanel')),
    coverage: [
      { symbol: 'useSPFxProperties', kind: 'write', notes: 'Updates SPFx properties through the provider.' },
      { symbol: 'useSPFxSessionStorage', kind: 'write', notes: 'Instance-scoped session storage.' },
      { symbol: 'useSPFxLocalStorage', kind: 'write', notes: 'Instance-scoped local storage.' },
      { symbol: 'useSPFxLogger', kind: 'write', notes: 'SPFx Log wrapper.' },
      { symbol: 'useSPFxPerformance', kind: 'read', notes: 'Performance timing helper.' },
    ],
  },
  {
    key: 'clients',
    title: 'Clients',
    iconName: 'Cloud',
    description: 'SPFx HTTP clients and state-managed invoke helpers.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-clients' */ './panels/ClientsPanel')),
    coverage: [
      { symbol: 'useSPFxHttpClient', kind: 'read', notes: 'External HTTP call through SPFx HttpClient.' },
      { symbol: 'useSPFxSPHttpClient', kind: 'read', notes: 'Current web REST read through SPHttpClient.' },
      { symbol: 'useSPFxMSGraphClient', kind: 'read', notes: 'Graph client status and optional /me call.' },
      { symbol: 'useSPFxAadHttpClient', kind: 'read', notes: 'AAD-secured client initialization by resource URL.' },
      { symbol: 'useSPFxAadTokenProvider', kind: 'read', notes: 'AAD token provider initialization status.' },
      { symbol: 'useSPFxApiPermissionPrecheck', kind: 'read', notes: 'Passive Graph/custom API delegated scope precheck.' },
    ],
  },
  {
    key: 'graph',
    title: 'Graph Data',
    iconName: 'OneDrive',
    description: 'Graph-backed convenience hooks for photo and OneDrive app data.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-graph-data' */ './panels/GraphDataPanel')),
    coverage: [
      { symbol: 'useSPFxUserPhoto', kind: 'read', notes: 'Current user photo with fallback.' },
      { symbol: 'useSPFxOneDriveAppData', kind: 'write', notes: 'App folder load/write with isolated demo file.' },
    ],
  },
  {
    key: 'site',
    title: 'Site',
    iconName: 'TableGroup',
    description: 'Collection-root key-value store shared by subsites, with manual CRUD.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-site' */ './panels/SitePanel')),
    coverage: [
      { symbol: 'useSPFxSiteKeyValueStore', kind: 'write', notes: 'Manual collection-root CRUD, safe text/JSON/bigint/date readback and conflict errors; disposable keys and root/subsite isolation.' },
    ],
  },
  {
    key: 'tenant',
    title: 'Tenant',
    iconName: 'Org',
    description: 'Tenant property reads and tenant key-value store CRUD.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-tenant' */ './panels/TenantPanel')),
    coverage: [
      { symbol: 'useSPFxTenantProperty', kind: 'read', notes: 'Read tenant properties from the app catalog.' },
      { symbol: 'useSPFxTenantKeyValueStore', kind: 'write', notes: 'Manual tenant CRUD, complete paging, effective permissions and observable text/JSON/bigint/date readback; reads never provision.' },
    ],
  },
  {
    key: 'configuration',
    title: 'Configuration',
    iconName: 'TableGroup',
    description: 'Tenant defaults, site overrides and explicitly scoped changes.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-configuration' */ './panels/ConfigurationPanel')),
    coverage: [
      { symbol: 'useSPFxConfigurationStore', kind: 'write', notes: 'Manual tenant/site/layered reads, provenance, physical inventories and explicit destination writes/removal with disposable keys.' },
    ],
  },
  {
    key: 'pnp',
    title: 'PnPjs',
    iconName: 'CloudDownload',
    description: 'PnP context, invoke/batch helpers, and list operations.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-pnp' */ './panels/PnPPanel')),
    coverage: [
      { symbol: 'useSPFxPnPContext', kind: 'read', notes: 'Current and optional cross-site SPFI contexts.' },
      { symbol: 'useSPFxPnP', kind: 'read', notes: 'PnP invoke and batching helper.' },
      { symbol: 'useSPFxPnPList', kind: 'write', notes: 'List query and optional CRUD operations.' },
      { symbol: 'useSPFxPnPListById', kind: 'read', notes: 'Manual first-page list query by GUID.' },
      { symbol: 'useSPFxPnPListByUrl', kind: 'read', notes: 'Manual first-page list query by decoded server-relative root URL.' },
      { symbol: 'useSPFxPnPListByPath', kind: 'read', notes: 'Manual first-page list query by decoded web-relative root path.' },
    ],
  },
  {
    key: 'search',
    title: 'Search',
    iconName: 'Search',
    description: 'SharePoint search, builder API, refiners, suggestions, and verticals.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-search' */ './panels/SearchPanel')),
    coverage: [
      { symbol: 'useSPFxPnPSearch', kind: 'read', notes: 'Search text, builder, refiners, and suggestions.' },
      { symbol: 'SearchVerticals', kind: 'snapshot', notes: 'Built-in source ids for common verticals.' },
    ],
  },
  {
    key: 'sharepoint-permission-trim', title: 'SharePoint Permission Trim', iconName: 'Permissions',
    description: 'Current-user access checks.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-sharepoint-permission-trim' */ './panels/SharePointPermissionTrimPanel')),
    coverage: [{ symbol: 'SharePointPermissionTrim', kind: 'read', notes: 'Actual native reads, all branches, selectors, protected state and explicit repeat; local transport fixtures and tenant validation are separate.' }],
  },
  {
    key: 'sharepoint-group-trim', title: 'SharePoint Group Trim', iconName: 'Permissions',
    description: 'Current-user access checks.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-sharepoint-group-trim' */ './panels/SharePointGroupTrimPanel')),
    coverage: [{ symbol: 'SharePointGroupTrim', kind: 'read', notes: 'Actual native reads, all branches, selectors, protected state and explicit repeat; local transport fixtures and tenant validation are separate.' }],
  },
  {
    key: 'entra-group-trim', title: 'Entra Group Trim', iconName: 'Permissions',
    description: 'Current-user access checks.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-entra-group-trim' */ './panels/EntraGroupTrimPanel')),
    coverage: [{ symbol: 'EntraGroupTrim', kind: 'read', notes: 'Actual native reads, all branches, selectors, protected state and explicit repeat; local transport fixtures and tenant validation are separate.' }],
  },
];

export const demoCoverage: readonly DemoCoverageItem[] = demoRegistry.flatMap(entry => entry.coverage);
