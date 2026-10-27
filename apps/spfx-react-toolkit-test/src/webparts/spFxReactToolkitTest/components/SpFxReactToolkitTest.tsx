import * as React from 'react';
import { MessageBar, MessageBarBody, Spinner, Tab, TabList, useId } from '@fluentui/react-components';
import styles from './SpFxReactToolkitTest.module.scss';
import { useSPFxTeams } from '@apvee/spfx-react-toolkit';
import {
  demoCoverage,
  demoRegistry,
  type DemoCoverageKind,
  type DemoRegistryEntry,
} from './demoRegistry';
import { DemoCard, DemoIcon, InfoGrid, StatusBadge, useDemoStyles } from './shared';

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
      title="Package Coverage"
      iconName="CompletedSolid"
      description="Every public provider and hook covered by the package sample surface is tracked here and verified by scripts/verify-example-coverage.cjs."
      contentPlacement="configuration"
      status={<div className={demoStyles.actions}>
        <StatusBadge label={`${symbols.size} symbols covered`} available={true} />
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

const DynamicDemoPanel: React.FC<{ entry: DemoRegistryEntry }> = ({ entry }) => {
  const demoStyles = useDemoStyles();
  const Component = entry.component;

  return (
    <div className={demoStyles.panel}>
      <MessageBar intent="info"><MessageBarBody>{entry.description}</MessageBarBody></MessageBar>
      <React.Suspense
        fallback={(
          <Spinner
            size="medium"
            label={`Loading ${entry.title} demos...`}
          />
        )}
      >
        <Component />
      </React.Suspense>
    </div>
  );
};

const SpFxReactToolkitTest: React.FC = () => {
  const { supported: hasTeamsContext } = useSPFxTeams();
  const demoStyles = useDemoStyles();
  const navigationId = useId('demo-navigation');
  const [selectedKey, setSelectedKey] = React.useState(demoRegistry[0].key);
  const activeEntry = demoRegistry.find(entry => entry.key === selectedKey) ?? demoRegistry[0];

  return (
    <section className={`${styles.spFxReactToolkitTest} ${demoStyles.root} ${hasTeamsContext ? styles.teams : ''}`}>
      <CoverageOverview />
      <TabList aria-label="SPFx React Toolkit Demo Tabs" className={demoStyles.tabList}
        selectTabOnFocus selectedValue={selectedKey} onTabSelect={(_event, data) => {
          if (typeof data.value === 'string' && demoRegistry.some(entry => entry.key === data.value)) setSelectedKey(data.value);
        }}>
        {demoRegistry.map(entry => (
          <Tab key={entry.key} value={entry.key} id={`${navigationId}-tab-${entry.key}`}
            aria-controls={`${navigationId}-panel-${entry.key}`} icon={<DemoIcon name={entry.iconName} />}>
            {entry.title}
          </Tab>
        ))}
      </TabList>
      <div role="tabpanel" id={`${navigationId}-panel-${activeEntry.key}`} aria-labelledby={`${navigationId}-tab-${activeEntry.key}`} tabIndex={0}>
        {/* Preserve Pivot's active-only mount cycle; host updates retain the selected component. */}
        <DynamicDemoPanel key={activeEntry.key} entry={activeEntry} />
      </div>
    </section>
  );
};

export default SpFxReactToolkitTest;
