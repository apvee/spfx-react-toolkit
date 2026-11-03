import * as React from 'react';
import { MessageBar, MessageBarBody, Spinner } from '@fluentui/react-components';
import type { ExecutableExampleEntry, ExecutableExampleProps, ExecutableExampleRegions } from '../generated/jsdoc-examples/catalogue';
import { DemoCard } from './DemoCard';
import { NativePropertyPaneContext, NativePropertyPaneDiagnostics } from './NativePropertyPaneContext';
import { useExecutableShellPort } from './useExecutableShellPort';

// Lazy component identities are shared across renders, resets and host presentation updates.
const loaders = new WeakMap<ExecutableExampleEntry, React.LazyExoticComponent<React.ComponentType<ExecutableExampleProps>>>();
function getComponent(entry: ExecutableExampleEntry): React.LazyExoticComponent<React.ComponentType<ExecutableExampleProps>> {
  let component = loaders.get(entry);
  if (!component) {
    component = React.lazy(() => entry.load().then(module => ({ default: module.Demo })));
    loaders.set(entry, component);
  }
  return component;
}

class ExampleLoadBoundary extends React.Component<{ readonly children?: React.ReactNode }, { readonly error?: Error }> {
  public state: { readonly error?: Error } = {};
  public static getDerivedStateFromError(error: Error): { readonly error: Error } { return { error }; }
  public render(): React.ReactNode {
    return this.state.error
      ? <MessageBar intent="error"><MessageBarBody>Impossibile caricare lo scenario: {this.state.error.message}</MessageBarBody></MessageBar>
      : this.props.children;
  }
}

interface ScenarioProps {
  readonly entry: ExecutableExampleEntry;
  readonly resetRevision: number;
  readonly render: (regions: ExecutableExampleRegions) => React.ReactElement;
}

function OrdinaryScenario({ entry, resetRevision, render }: ScenarioProps): React.ReactElement {
  const Demo = getComponent(entry);
  return <Demo key={entry.id} render={render} resetRevision={resetRevision} />;
}

/** Separate component keeps the private Shell hook unconditional and out of ordinary examples. */
function ShellScenario({ entry, resetRevision, render }: ScenarioProps): React.ReactElement {
  const shell = useExecutableShellPort();
  const Demo = getComponent(entry);
  return <Demo key={entry.id} render={render} resetRevision={resetRevision} shell={shell} />;
}

/** Native port is optional and instance owned above Shell body and scenario selection. */
function NativeScenario({ entry, resetRevision, render }: ScenarioProps): React.ReactElement {
  const host = React.useContext(NativePropertyPaneContext);
  const Demo = getComponent(entry);
  const nativeRender = React.useCallback((regions: ExecutableExampleRegions): React.ReactElement => render({ ...regions,
    diagnostics: <>{regions.diagnostics}<NativePropertyPaneDiagnostics /></>
  }), [render]);
  return <Demo key={entry.id} render={nativeRender} resetRevision={resetRevision} {...(host ? { propertyPane: host.port } : {})} />;
}

export function ExecutableExampleHost({ entry, resetRevision }: {
  readonly entry: ExecutableExampleEntry;
  readonly resetRevision: number;
}): React.ReactElement {
  const render = React.useCallback((regions: ExecutableExampleRegions): React.ReactElement => (
    <DemoCard title={entry.title} iconName="Info" contentPlacement={entry.metadata.contentPlacement} {...regions} />
  ), [entry]);
  const Scenario = entry.metadata.hostAdapter === 'webPartShell' ? ShellScenario : entry.metadata.hostAdapter === 'propertyPane' ? NativeScenario : OrdinaryScenario;
  return <ExampleLoadBoundary key={entry.id}>
    <React.Suspense fallback={<Spinner size="medium" label={`Caricamento ${entry.title}...`} />}>
      <Scenario entry={entry} resetRevision={resetRevision} render={render} />
    </React.Suspense>
  </ExampleLoadBoundary>;
}
