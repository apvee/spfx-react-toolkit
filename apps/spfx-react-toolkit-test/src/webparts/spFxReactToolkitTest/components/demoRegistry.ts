import * as React from 'react';

import type { DemoCoverageItem, DemoRegistryEntry } from './demoRegistry.types';
export type { DemoCoverageKind, DemoCoverageItem, DemoRegistryEntry } from './demoRegistry.types';

const lazyPanel = (
  loader: () => Promise<{ default: React.ComponentType }>
): React.LazyExoticComponent<React.ComponentType> => React.lazy(loader);

/** Temporary panels only: these entries are not generated features or executed test results. */
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

];

export const demoCoverage: readonly DemoCoverageItem[] = demoRegistry.flatMap(entry => entry.coverage);
