import * as React from 'react';
import { Button, Field, MessageBar, MessageBarBody, Select, Spinner, Text } from '@fluentui/react-components';
import type { ExecutableGroup } from './generated/jsdoc-examples/catalogue';
import type { DemoRegistryEntry } from './demoRegistry';
import { ExampleExplorerNavigation, getNavigationGroups } from './shared/ExampleExplorerNavigation';
import { useExampleExplorerState } from './shared/ExampleExplorerState';
import { ExecutableExampleHost } from './shared/ExecutableExampleHost';
import { useExampleExplorerStyles } from './shared/exampleExplorerStyles';

export function ExampleExplorer({ groups, legacyEntries }: {
  readonly groups: readonly ExecutableGroup[];
  readonly legacyEntries: readonly DemoRegistryEntry[];
}): React.ReactElement {
  const state = useExampleExplorerState();
  const styles = useExampleExplorerStyles();
  const navigationGroups = getNavigationGroups(groups, legacyEntries);
  const features = navigationGroups.flatMap(group => group.features);
  const selected = features.find(feature => feature.key === state.selectedFeatureKey) ?? features[0];
  const group = navigationGroups.find(item => item.features.some(feature => feature.key === selected?.key));
  const feature = groups.flatMap(item => item.features).find(item => item.key === selected?.key);
  const entry = feature?.scenarios.find(item => item.id === state.selectedScenarioByFeature[feature.key]) ?? feature?.scenarios[0];
  const legacyEntry = legacyEntries.find(item => `legacy:${item.key}` === selected?.key);
  const LegacyPanel = legacyEntry?.component;
  return <div className={styles.root}>
    <div className={styles.layout}>
      <ExampleExplorerNavigation groups={groups} legacyEntries={legacyEntries} />
      <div className={styles.panel} role="region" aria-label={selected?.title ?? 'Esempi'}>
        {selected ? <>
          <Text className={styles.breadcrumb}>{group?.title} / {selected.title}</Text>
          <Text as="h2" size={600} weight="semibold" className={styles.heading}>{selected.title}</Text>
          {entry && <>
            <div className={styles.toolbar}>
              {feature && feature.scenarios.length > 1 && <Field label="Scenario">
                <Select aria-label="Scenario" value={entry.id} onChange={(_event, data) => state.selectScenario(feature.key, data.value)}>
                  {feature.scenarios.map(scenario => <option key={scenario.id} value={scenario.id}>{scenario.title}</option>)}
                </Select>
              </Field>}
              <Button onClick={state.resetScenario}>Ripristina scenario</Button>
            </div>
            <Text><strong>Cosa osservare: </strong>{entry.metadata.observe}</Text>
            <ExecutableExampleHost entry={entry} resetRevision={state.resetRevision} />
          </>}
          {legacyEntry && LegacyPanel && <>
            <MessageBar intent="info"><MessageBarBody>{legacyEntry.description}</MessageBarBody></MessageBar>
            <React.Suspense fallback={<Spinner size="medium" label={`Caricamento ${legacyEntry.title}...`} />}>
              <LegacyPanel key={legacyEntry.key} />
            </React.Suspense>
          </>}
        </> : <Text>Nessuno scenario disponibile.</Text>}
      </div>
    </div>
  </div>;
}
