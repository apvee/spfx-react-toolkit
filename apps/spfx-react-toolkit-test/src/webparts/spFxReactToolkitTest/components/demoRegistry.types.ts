import type * as React from 'react';

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
