import * as React from 'react';
import { Button, Field, Select, Text, useId } from '@fluentui/react-components';
import { ChevronDownRegular, ChevronRightRegular } from '@fluentui/react-icons';
import type { ExecutableGroup } from '../generated/jsdoc-examples/catalogue';
import type { DemoRegistryEntry } from '../demoRegistry.types';
import { useExampleExplorerState } from './ExampleExplorerState';
import { useExampleExplorerStyles } from './exampleExplorerStyles';

interface NavigationGroup {
  readonly key: string;
  readonly title: string;
  readonly features: readonly { readonly key: string; readonly title: string }[];
}

/** The temporary bridge keeps one entry per existing panel, regardless of its covered symbols. */
export function getNavigationGroups(groups: readonly ExecutableGroup[], legacyEntries: readonly DemoRegistryEntry[]): readonly NavigationGroup[] {
  const generated = groups.map(group => ({ ...group, features: group.features.filter(feature => feature.scenarios.length > 0) })).filter(group => group.features.length > 0);
  return [...generated, ...(legacyEntries.length > 0 ? [{ key: 'legacy', title: 'Scenari esistenti',
    features: legacyEntries.map(entry => ({ key: `legacy:${entry.key}`, title: entry.title })) }] : [])];
}

export function ExampleExplorerNavigation({ groups, legacyEntries }: {
  readonly groups: readonly ExecutableGroup[];
  readonly legacyEntries: readonly DemoRegistryEntry[];
}): React.ReactElement {
  const state = useExampleExplorerState();
  const styles = useExampleExplorerStyles();
  const navigationId = useId('example-navigation');
  const navigationGroups = getNavigationGroups(groups, legacyEntries);
  const selectedKey = navigationGroups.flatMap(group => group.features).find(feature => feature.key === state.selectedFeatureKey)?.key ?? navigationGroups[0]?.features[0]?.key;
  return <>
    <Field label="Funzionalità" className={styles.mobile}>
      <Select className={styles.selectRoot} select={{ className: styles.selectInput }} aria-label="Funzionalità" value={selectedKey ?? ''} disabled={!selectedKey} onChange={(_event, data) => state.selectFeature(data.value)}>
        {navigationGroups.map(group => <optgroup key={group.key} label={group.title}>
          {group.features.map(feature => <option key={feature.key} value={feature.key}>{feature.title}</option>)}
        </optgroup>)}
      </Select>
    </Field>
    <nav aria-label="Gruppi e funzionalità" className={styles.sidebar}>
      {navigationGroups.map(group => {
        const open = state.expandedGroupKeys.includes(group.key);
        const selected = group.features.find(feature => feature.key === selectedKey);
        const panelId = `${navigationId}-${group.key}`;
        return <div key={group.key} className={styles.group}>
          <Button className={styles.groupHeader} appearance="subtle" icon={open ? <ChevronDownRegular /> : <ChevronRightRegular />}
            aria-expanded={open} aria-controls={panelId} onClick={() => state.toggleGroup(group.key)}>
            <span className={styles.groupLabel}>
              <span className={styles.truncated} title={group.title}>{group.title}</span>
              {!open && selected && <Text size={200} className={styles.truncated} title={selected.title}>{selected.title}</Text>}
            </span>
          </Button>
          <div id={panelId} hidden={!open}>
            <div className={styles.featureList}>
              {group.features.map(feature => <Button key={feature.key} className={styles.feature} title={feature.title}
                appearance={feature.key === selectedKey ? 'secondary' : 'subtle'} aria-current={feature.key === selectedKey ? 'true' : undefined}
                data-example-feature={feature.key} onClick={() => state.selectFeature(feature.key)}>
                <span className={styles.truncated}>{feature.title}</span>
              </Button>)}
            </div>
          </div>
        </div>;
      })}
    </nav>
  </>;
}
