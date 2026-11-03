import * as React from 'react';
import { Text } from '@fluentui/react-components';
import styles from './SpFxReactToolkitTest.module.scss';
import { ExampleExplorer } from './ExampleExplorer';
import { webPartExecutableGroups } from './shared/webPartExecutableGroups';
import { useSPFxTeams } from '@apvee/spfx-react-toolkit';
import { DemoCard, InfoGrid, useDemoStyles } from './shared';

function CoverageOverview(): React.ReactElement {
  const features = webPartExecutableGroups.flatMap(group => group.features);
  const scenarios = features.flatMap(feature => feature.scenarios);
  return <DemoCard title="Disponibilità degli esempi" iconName="CompletedSolid"
    description="Disponibilità generata per questo host WebPart. Questi conteggi non attestano esecuzione, provider montati o superamento di prove."
    contentPlacement="configuration"
    status={<Text>{features.length} funzionalità generate · {scenarios.length} scenari disponibili per WebPart</Text>}
    diagnosticsLabel="Detailed package coverage"
    diagnostics={<InfoGrid rows={webPartExecutableGroups.map(group => ({ label: group.title,
      value: `${group.features.length} funzionalità · ${group.features.reduce((count, feature) => count + feature.scenarios.length, 0)} scenari`, icon: 'Info' }))} />} />;
}

const SpFxReactToolkitTest: React.FC = () => {
  const { supported: hasTeamsContext } = useSPFxTeams();
  const demoStyles = useDemoStyles();
  return (
    <section className={`${styles.spFxReactToolkitTest} ${demoStyles.root} ${hasTeamsContext ? styles.teams : ''}`}>
      <ExampleExplorer groups={webPartExecutableGroups} />
      <CoverageOverview />
    </section>
  );
};

export default SpFxReactToolkitTest;
