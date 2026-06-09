import * as React from 'react';

export type DemoCoverageKind = 'host' | 'snapshot' | 'read' | 'write';

export interface DemoCoverageItem {
  readonly symbol: string;
  readonly kind: DemoCoverageKind;
  readonly notes: string;
}

export interface DemoRegistryEntry {
  readonly key: string;
  readonly title: string;
  readonly iconName: string;
  readonly description: string;
  readonly coverage: readonly DemoCoverageItem[];
  readonly component: React.LazyExoticComponent<React.ComponentType>;
}

const lazyPanel = (
  loader: () => Promise<{ default: React.ComponentType }>
): React.LazyExoticComponent<React.ComponentType> => React.lazy(loader);

export const demoRegistry: readonly DemoRegistryEntry[] = [
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
    key: 'tenant',
    title: 'Tenant',
    iconName: 'Org',
    description: 'Tenant property reads and tenant key-value store CRUD.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-tenant' */ './panels/TenantPanel')),
    coverage: [
      { symbol: 'useSPFxTenantProperty', kind: 'read', notes: 'Read tenant properties from the app catalog.' },
      { symbol: 'useSPFxTenantKeyValueStore', kind: 'write', notes: 'REST-based tenant key-value store CRUD.' },
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
];

export const demoCoverage: readonly DemoCoverageItem[] = demoRegistry.flatMap(entry => entry.coverage);
