import * as React from 'react';
import {
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  Spinner,
  SpinnerSize,
  Stack,
} from '@fluentui/react';
import styles from './SpFxReactToolkitTest.module.scss';
import { useSPFxTeams } from '../../../hooks';
import {
  demoCoverage,
  demoRegistry,
  type DemoCoverageKind,
  type DemoRegistryEntry,
} from './demoRegistry';
import { DemoCard, InfoGrid, StatusBadge } from './shared';

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
  const symbols = new Set(demoCoverage.map(item => item.symbol));

  return (
    <DemoCard
      title="Package Coverage"
      iconName="CompletedSolid"
      description="Every public provider and hook covered by the package sample surface is tracked here and verified by scripts/verify-example-coverage.cjs."
    >
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <StatusBadge label={`${symbols.size} symbols covered`} available={true} />
        <StatusBadge label={`${countByKind('host')} host providers`} available={true} />
        <StatusBadge label={`${countByKind('snapshot')} snapshots`} available={true} />
        <StatusBadge label={`${countByKind('read')} read demos`} available={true} />
        <StatusBadge label={`${countByKind('write')} write demos`} available={true} />
      </Stack>
      <InfoGrid
        rows={demoRegistry.map(entry => ({
          label: entry.title,
          value: `${entry.coverage.length} item(s): ${entry.coverage.map(item => `${item.symbol} (${coverageKindLabels[item.kind]})`).join(', ')}`,
          icon: entry.iconName,
        }))}
      />
    </DemoCard>
  );
};

const DynamicDemoPanel: React.FC<{ entry: DemoRegistryEntry }> = ({ entry }) => {
  const Component = entry.component;

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16 } }}>
      <MessageBar messageBarType={MessageBarType.info}>
        {entry.description}
      </MessageBar>
      <React.Suspense
        fallback={(
          <Spinner
            size={SpinnerSize.medium}
            label={`Loading ${entry.title} demos...`}
          />
        )}
      >
        <Component />
      </React.Suspense>
    </Stack>
  );
};

const SpFxReactToolkitTest: React.FC = () => {
  const { supported: hasTeamsContext } = useSPFxTeams();

  return (
    <section className={`${styles.spFxReactToolkitTest} ${hasTeamsContext ? styles.teams : ''}`}>
      <CoverageOverview />
      <Pivot aria-label="SPFx React Toolkit Demo Tabs" style={{ marginTop: 20 }}>
        {demoRegistry.map(entry => (
          <PivotItem key={entry.key} headerText={entry.title} itemIcon={entry.iconName}>
            <DynamicDemoPanel entry={entry} />
          </PivotItem>
        ))}
      </Pivot>
    </section>
  );
};

export default SpFxReactToolkitTest;
