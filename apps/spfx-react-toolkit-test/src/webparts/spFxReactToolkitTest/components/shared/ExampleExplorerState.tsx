import * as React from 'react';
import type { ExecutableGroup } from '../generated/jsdoc-examples/catalogue';

export interface ExampleExplorerState {
  readonly selectedFeatureKey?: string;
  readonly selectedScenarioByFeature: Readonly<Record<string, string>>;
  readonly expandedGroupKeys: readonly string[];
  readonly resetRevision: number;
  selectFeature(featureKey: string): void;
  selectScenario(featureKey: string, scenarioId: string): void;
  toggleGroup(groupKey: string): void;
  resetScenario(): void;
}

const ExplorerContext = React.createContext<ExampleExplorerState | undefined>(undefined);

/** Navigation survives the Shell's real subtle/Card body boundary; scenario state does not. */
export function ExampleExplorerStateProvider({ initialGroups, children }: {
  readonly initialGroups: readonly ExecutableGroup[];
  readonly children?: React.ReactNode;
}): React.ReactElement {
  const groups = initialGroups.filter(group => group.features.some(feature => feature.scenarios.length > 0));
  const firstFeature = groups[0]?.features.find(feature => feature.scenarios.length > 0);
  const [selection, setSelection] = React.useState<string | undefined>(() => firstFeature?.key);
  const [selectedScenarioByFeature, setScenarios] = React.useState<Readonly<Record<string, string>>>({});
  const [expandedGroupKeys, setExpanded] = React.useState<readonly string[]>(() => groups[0] ? [groups[0].key] : ['legacy']);
  const [resetRevision, setResetRevision] = React.useState(0);
  const selectedFeatureKey = selection?.startsWith('legacy:') || groups.some(group => group.features.some(feature => feature.key === selection && feature.scenarios.length > 0))
    ? selection : firstFeature?.key;

  const selectFeature = React.useCallback((featureKey: string): void => {
    const groupKey = featureKey.startsWith('legacy:') ? 'legacy' : initialGroups.find(group => group.features.some(feature => feature.key === featureKey && feature.scenarios.length > 0))?.key;
    if (!groupKey) return;
    setSelection(featureKey);
    setExpanded(previous => previous.includes(groupKey) ? previous : [...previous, groupKey]);
  }, [initialGroups]);
  const selectScenario = React.useCallback((featureKey: string, scenarioId: string): void => {
    const feature = initialGroups.flatMap(group => group.features).find(item => item.key === featureKey);
    if (!feature?.scenarios.some(entry => entry.id === scenarioId)) return;
    setScenarios(previous => ({ ...previous, [featureKey]: scenarioId }));
  }, [initialGroups]);
  const toggleGroup = React.useCallback((groupKey: string): void => {
    setExpanded(previous => previous.includes(groupKey) ? previous.filter(key => key !== groupKey) : [...previous, groupKey]);
  }, []);
  const resetScenario = React.useCallback((): void => setResetRevision(previous => previous + 1), []);
  const value = React.useMemo(() => ({ selectedFeatureKey, selectedScenarioByFeature, expandedGroupKeys, resetRevision,
    selectFeature, selectScenario, toggleGroup, resetScenario }),
  [selectedFeatureKey, selectedScenarioByFeature, expandedGroupKeys, resetRevision, selectFeature, selectScenario, toggleGroup, resetScenario]);
  return <ExplorerContext.Provider value={value}>{children}</ExplorerContext.Provider>;
}

export function useExampleExplorerState(): ExampleExplorerState {
  const state = React.useContext(ExplorerContext);
  if (!state) throw new Error('ExampleExplorer requires ExampleExplorerStateProvider.');
  return state;
}
