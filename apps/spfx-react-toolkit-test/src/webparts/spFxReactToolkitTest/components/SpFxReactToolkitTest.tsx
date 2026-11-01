import * as React from 'react';
import { Text } from '@fluentui/react-components';
import styles from './SpFxReactToolkitTest.module.scss';
import { ExampleExplorer } from './ExampleExplorer';
import { executableExamples, executableGroups } from './generated/jsdoc-examples/catalogue';
import { useSPFxTeams } from '@apvee/spfx-react-toolkit';
import {
  demoCoverage,
  demoRegistry,
  type DemoCoverageKind,
} from './demoRegistry';
import { DemoCard, InfoGrid, StatusBadge, useDemoStyles } from './shared';

const coverageKindLabels: Record<DemoCoverageKind, string> = {
  host: 'Host',
  snapshot: 'Snapshot',
  read: 'Read',
  write: 'Write',
};

function countByKind(kind: DemoCoverageKind): number {
  return demoCoverage.filter(item => item.kind === kind).length;
}

const CoverageOverview: React.FC = () => {
  const demoStyles = useDemoStyles();
  const symbols = new Set(demoCoverage.map(item => item.symbol));

  return (
    <DemoCard
      title="Disponibilità degli esempi"
      iconName="CompletedSolid"
      description="Gli scenari generati e i pannelli transitori sono conteggi distinti; la disponibilità non attesta l’esecuzione o il superamento di prove."
      contentPlacement="configuration"
      status={<div className={demoStyles.actions}>
        <Text>{executableGroups.reduce((count, group) => count + group.features.length, 0)} funzionalità generate · {executableExamples.length} scenari generati · {demoRegistry.length} pannelli transitori</Text>
        <StatusBadge label={`${symbols.size} simboli nel registro transitorio`} available={true} />
        <StatusBadge label={`${countByKind('host')} host providers`} available={true} />
        <StatusBadge label={`${countByKind('snapshot')} snapshots`} available={true} />
        <StatusBadge label={`${countByKind('read')} read demos`} available={true} />
        <StatusBadge label={`${countByKind('write')} write demos`} available={true} />
      </div>}
      diagnosticsLabel="Detailed package coverage"
      diagnostics={<InfoGrid
        rows={demoRegistry.map(entry => ({
          label: entry.title,
          value: `${entry.coverage.length} item(s): ${entry.coverage.map(item => `${item.symbol} (${coverageKindLabels[item.kind]})`).join(', ')}`,
          icon: entry.iconName,
        }))}
      />}
    />
  );
};

const SpFxReactToolkitTest: React.FC = () => {
  const { supported: hasTeamsContext } = useSPFxTeams();
  const demoStyles = useDemoStyles();
  return (
    <section className={`${styles.spFxReactToolkitTest} ${demoStyles.root} ${hasTeamsContext ? styles.teams : ''}`}>
      <ExampleExplorer groups={executableGroups} legacyEntries={demoRegistry} />
      <CoverageOverview />
    </section>
  );
};

export default SpFxReactToolkitTest;
