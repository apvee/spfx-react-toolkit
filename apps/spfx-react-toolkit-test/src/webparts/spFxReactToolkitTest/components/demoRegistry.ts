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
