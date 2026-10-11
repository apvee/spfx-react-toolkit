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
    key: 'web-part-shell',
    title: 'WebPartShell',
    iconName: 'Page',
    description: 'One root Shell, editable rich title, persistent children and host presentation overrides.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-web-part-shell' */ './panels/WebPartShellPanel')),
    coverage: [
      { symbol: 'WebPartShell', kind: 'host', notes: 'Actual root wrapper, hook action, title visibility, rich/plain/empty, child counter/note, theme/RTL and surface/padding.' },
      { symbol: 'getWebPartShellPropertiesMetadata', kind: 'snapshot', notes: 'Web part metadata getter marks the plain title as searchable.' },
    ],
  },
  {
    key: 'data-grid',
    title: 'DataGrid',
    iconName: 'Table',
    description: 'Sixteen cell presentations, stable IDs and consumer-owned rows, sort intent and selection.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-data-grid' */ './panels/DataGridPanel')),
    coverage: [
      { symbol: 'DataGrid', kind: 'snapshot', notes: 'Dynamic valid/missing/invalid values, native cells, absolute dates and localized statuses.' },
      { symbol: 'createDataGrid', kind: 'snapshot', notes: 'Typed sixteen-column family, descriptors, controlled selection/sort, activation, loading, resizing and immutable Charts updates.' },
    ],
  },
  {
    key: 'form',
    title: 'Form',
    iconName: 'Edit',
    description: 'Typed live values, settings field rows, all adapters, nested scopes, arrays and save/reset ownership.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-form' */ './panels/FormPanel')),
    coverage: [
      { symbol: 'Form', kind: 'snapshot', notes: 'Dynamic Form.Root/Input/Submit/Reset alongside the typed complete-family record editor.' },
      { symbol: 'createForm', kind: 'snapshot', notes: '22 adapters, automatic SettingsGroup field labels/hints, suite hooks, four compositions, rules/schema, native files, drafts, reset, pending save and retries.' },
    ],
  },
  {
    key: 'stage',
    title: 'Stage',
    iconName: 'OpenPane',
    description: 'Explicit Dialog, Drawer and Popover surfaces with shared content and guarded forms.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-stage' */ './panels/StagePanel')),
    coverage: [{ symbol: 'Stage', kind: 'snapshot', notes: 'Three presentations, sizes, literal/JSX headers, scroll, retention, nested popups, guarded Form save/reset, themes/RTL and selectable Card isolation.' }],
  },
  {
    key: 'groups',
    title: 'SettingsGroup and AccordionGroup',
    iconName: 'Settings',
    description: 'Caller-owned settings and controlled disclosure with retained panel state.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-groups' */ './panels/GroupsPanel')),
    coverage: [
      { symbol: 'SettingsGroup', kind: 'snapshot', notes: 'Caller-labeled Input/Switch rows, literal text/JSX, actions, width, theme, direction, hidden and root ref; automatic field composition also runs in Form.' },
      { symbol: 'AccordionGroup', kind: 'snapshot', notes: 'Controlled openItems, multiple, disabled, mounted note persistence, unmount reset and root ref.' },
    ],
  },
  {
    key: 'state-message',
    title: 'StateMessage',
    iconName: 'Info',
    description: 'Section states, theme-aware artwork, container sizing and accessible actions.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-state-message' */ './panels/StateMessagePanel')),
    coverage: [{ symbol: 'StateMessage', kind: 'snapshot', notes: 'Five presets, light/dark, RTL, responsive host, custom/hidden visual, heading, role, button/link actions and root ref.' }],
  },
  {
    key: 'auto-skeleton',
    title: 'AutoSkeleton',
    iconName: 'Processing',
    description: 'Measured loading geometry, persistent child state, resize, focus and direction.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-auto-skeleton' */ './panels/AutoSkeletonPanel')),
    coverage: [
      { symbol: 'AutoSkeleton', kind: 'snapshot', notes: 'Loading, animation, minimum size, ignored subtrees, empty content, resize, themes/RTL and native ref/focus/state observations.' },
    ],
  },
  {
    key: 'relative-time',
    title: 'RelativeTime',
    iconName: 'Clock',
    description: 'Explicit instants, calendar boundaries, clock modes and accessible full dates.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-relative-time' */ './panels/RelativeTimePanel')),
    coverage: [
      { symbol: 'RelativeTime', kind: 'snapshot', notes: 'Editable props, fixed/live/frozen clock, tooltip, fallback, native ref/click and FormattedText/Badge composition.' },
    ],
  },
  {
    key: 'formatted-text',
    title: 'FormattedText',
    iconName: 'Font',
    description: 'Inline Markdown, typography, native elements and constrained overflow.',
    component: lazyPanel(() => import(/* webpackChunkName: 'spfx-demo-formatted-text' */ './panels/FormattedTextPanel')),
    coverage: [
      { symbol: 'FormattedText', kind: 'snapshot', notes: 'Editable props, narrow width, RTL, mixed JSX and observable native ref/click behavior.' },
    ],
  },
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
