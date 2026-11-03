import type { ExecutableExampleEntry, ExecutableGroup } from '../generated/jsdoc-examples/catalogue';

export type ExecutableHostKind = 'webPart' | 'applicationCustomizer';

/** Consumer policy only; canonical metadata remains the sole scenario inventory. */
export function isExecutableEntryForHost(entry: ExecutableExampleEntry, host: ExecutableHostKind | undefined): boolean {
  if (host === 'applicationCustomizer') return entry.metadata.host === 'applicationCustomizer' && entry.metadata.hostAdapter === 'applicationCustomizerTop';
  if (entry.metadata.host === 'none') return entry.metadata.hostAdapter === undefined;
  return host === 'webPart' && entry.metadata.host === 'webPart' &&
    (entry.metadata.hostAdapter === undefined || entry.metadata.hostAdapter === 'webPartShell' || entry.metadata.hostAdapter === 'propertyPane');
}

/** Filter before selection/counting, retaining the exact canonical lazy entry objects. */
export function selectExecutableGroupsForHost(groups: readonly ExecutableGroup[], host: ExecutableHostKind): readonly ExecutableGroup[] {
  return groups.map(group => ({ ...group, features: group.features.map(feature => ({ ...feature,
    scenarios: feature.scenarios.filter(entry => isExecutableEntryForHost(entry, host))
  })).filter(feature => feature.scenarios.length > 0) })).filter(group => group.features.length > 0);
}
