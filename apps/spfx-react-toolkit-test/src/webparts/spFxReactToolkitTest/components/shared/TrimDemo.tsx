import * as React from 'react';
import { Button } from '@fluentui/react-components';

export function TrimProtectedContent({ onMount, onUnmount }: { onMount: () => void; onUnmount: () => void }): React.ReactElement {
  const [count, setCount] = React.useState(0);
  React.useEffect(() => { onMount(); return onUnmount; }, [onMount, onUnmount]);
  return <div data-trim-demo="protected"><p>Allowed: protected content is mounted.</p><Button data-trim-demo="increment" onClick={() => setCount(value => value + 1)}>Protected counter: {count}</Button></div>;
}
export function useTrimObservation(): { onMount: () => void; onUnmount: () => void; gauge: React.ReactElement } {
  const [mounts, setMounts] = React.useState(0);
  const [active, setActive] = React.useState(false);
  const onMount = React.useCallback(() => { setMounts(value => value + 1); setActive(true); }, []);
  const onUnmount = React.useCallback(() => setActive(false), []);
  return { onMount, onUnmount, gauge: <p data-trim-demo="mount-gauge">Protected mounts: {mounts}; active: {String(active)}</p> };
}
export function TrimBranchStatus({ state }: { state: 'loading' | 'denied' | 'error' }): React.ReactElement {
  return <p data-trim-demo="state" data-trim-state={state}>{state === 'loading' ? 'Loading: checking current access.' : state === 'denied' ? 'Denied: the requested access was not observed.' : 'Error: access could not be checked. Verify selectors, consent and transport.'}</p>;
}
