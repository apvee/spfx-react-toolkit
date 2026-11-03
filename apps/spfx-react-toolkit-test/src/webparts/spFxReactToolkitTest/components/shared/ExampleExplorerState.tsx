import * as React from 'react';
import type { ExecutableGroup } from '../generated/jsdoc-examples/catalogue';

interface ExampleExplorerState {
  readonly selectedFeatureKey?: string;
  readonly selectedScenarioByFeature: Readonly<Record<string, string>>;
  readonly expandedGroupKeys: readonly string[];
  readonly resetRevision: number;
  selectFeature(featureKey: string): void;
  selectScenario(featureKey: string, scenarioId: string): void;
  toggleGroup(groupKey: string): void;
  resetScenario(trigger?: HTMLButtonElement): void;
  readonly resetButtonRef: React.RefCallback<HTMLButtonElement>;
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
  const [expandedGroupKeys, setExpanded] = React.useState<readonly string[]>(() => groups[0] ? [groups[0].key] : []);
  const [resetRevision, setResetRevision] = React.useState(0);
  const resetFocus = React.useRef<{
    readonly trigger: HTMLButtonElement;
    readonly document: Document;
    readonly cancelOnFocus: (event: FocusEvent) => void;
    readonly cancelOnPointer: () => void;
  }>();
  const clearResetFocus = React.useCallback((): void => {
    const intent = resetFocus.current;
    if (!intent) return;
    intent.document.removeEventListener('focusin', intent.cancelOnFocus, true);
    intent.document.removeEventListener('pointerdown', intent.cancelOnPointer, true);
    resetFocus.current = undefined;
  }, []);
  React.useEffect(() => clearResetFocus, [clearResetFocus]);
  // React ref callbacks run during commit. Only an explicitly focused Reset
  // can authorize recovery when the real Shell replaces this toolbar.
  const resetButtonRef = React.useCallback<React.RefCallback<HTMLButtonElement>>(element => {
    const intent = resetFocus.current;
    if (!element || !intent || intent.trigger.isConnected) return;
    const active = intent.document.activeElement;
    const abandoned = active === intent.document.body || active === intent.document.documentElement;
    clearResetFocus();
    if (abandoned && element.ownerDocument === intent.document) element.focus();
  }, [clearResetFocus]);
  const selectedFeatureKey = groups.some(group => group.features.some(feature => feature.key === selection && feature.scenarios.length > 0))
    ? selection : firstFeature?.key;

  const selectFeature = React.useCallback((featureKey: string): void => {
    clearResetFocus();
    const groupKey = initialGroups.find(group => group.features.some(feature => feature.key === featureKey && feature.scenarios.length > 0))?.key;
    if (!groupKey) return;
    setSelection(featureKey);
    setExpanded(previous => previous.includes(groupKey) ? previous : [...previous, groupKey]);
  }, [initialGroups, clearResetFocus]);
  const selectScenario = React.useCallback((featureKey: string, scenarioId: string): void => {
    clearResetFocus();
    const feature = initialGroups.flatMap(group => group.features).find(item => item.key === featureKey);
    if (!feature?.scenarios.some(entry => entry.id === scenarioId)) return;
    setScenarios(previous => ({ ...previous, [featureKey]: scenarioId }));
  }, [initialGroups, clearResetFocus]);
  const toggleGroup = React.useCallback((groupKey: string): void => {
    setExpanded(previous => previous.includes(groupKey) ? previous.filter(key => key !== groupKey) : [...previous, groupKey]);
  }, []);
  const resetScenario = React.useCallback((trigger?: HTMLButtonElement): void => {
    clearResetFocus();
    if (trigger && trigger.ownerDocument.activeElement === trigger) {
      const document = trigger.ownerDocument;
      const cancelOnFocus = (event: FocusEvent): void => {
        if (event.target !== trigger) clearResetFocus();
      };
      const cancelOnPointer = (): void => clearResetFocus();
      resetFocus.current = { trigger, document, cancelOnFocus, cancelOnPointer };
      document.addEventListener('focusin', cancelOnFocus, true);
      document.addEventListener('pointerdown', cancelOnPointer, true);
    }
    setResetRevision(previous => previous + 1);
  }, [clearResetFocus]);
  const value = React.useMemo(() => ({ selectedFeatureKey, selectedScenarioByFeature, expandedGroupKeys, resetRevision,
    selectFeature, selectScenario, toggleGroup, resetScenario, resetButtonRef }),
  [selectedFeatureKey, selectedScenarioByFeature, expandedGroupKeys, resetRevision, selectFeature, selectScenario, toggleGroup, resetScenario, resetButtonRef]);
  return <ExplorerContext.Provider value={value}>{children}</ExplorerContext.Provider>;
}

export function useExampleExplorerState(): ExampleExplorerState {
  const state = React.useContext(ExplorerContext);
  if (!state) throw new Error('ExampleExplorer requires ExampleExplorerStateProvider.');
  return state;
}
