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
    key: 'imports',
    title: 'Imports',
    iconName: 'Code',
    description: 'Root, domain and legacy imports share callback behavior and style composition.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-imports' */ './panels/ImportsPanel')),
    coverage: [], // Existing React Hooks and Styles entries own the exported symbols.
  },
  {
    key: 'react-hooks',
    title: 'React Hooks',
    iconName: 'Code',
    description: 'Reusable React hooks that do not require an SPFx provider.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-react-hooks' */ './panels/ReactHooksPanel')),
    coverage: [
      { symbol: 'useStableCallback', kind: 'snapshot', notes: 'Retained callback identity and updated counter after re-render.' },
    ],
  },
  {
    key: 'styles',
    title: 'Styles',
    iconName: 'Color',
    description: 'Interactive descriptor catalog, themes, state composition and independent query regions.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-styles' */ './panels/StylesPanel')),
    coverage: [
      { symbol: 'useSx', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'width', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'minWidth', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'height', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'minHeight', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'maxWidth', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'maxHeight', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'flex', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'flexItem', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'grid', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'alignItems', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'justifyContent', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'alignSelf', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'gap', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'padding', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'paddingInline', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'paddingBlock', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'paddingInlineStart', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'paddingInlineEnd', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'paddingBlockStart', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'paddingBlockEnd', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'margin', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'marginInline', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'marginBlock', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'marginInlineStart', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'marginInlineEnd', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'marginBlockStart', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'marginBlockEnd', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'foreground', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'background', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'presets', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'typography', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'textAlign', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'text', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'borderWidth', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'borderStyle', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'borderColor', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'borderRadius', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'boxShadow', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'overflow', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'scrollbar', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'container', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'responsive', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'viewport', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'hover', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'active', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
      { symbol: 'focusVisible', kind: 'snapshot', notes: 'Interactive Styles controls apply real descriptors; local browser fixture complements host validation.' },
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
      { symbol: 'useSPFxSiteKeyValueStore', kind: 'write', notes: 'Manual collection-root CRUD using disposable site demo keys.' },
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
];

export const demoCoverage: readonly DemoCoverageItem[] = demoRegistry.flatMap(entry => entry.coverage);
