/**
 * Generic React callback utilities with no SPFx provider requirement.
 *
 * @packageDocumentation
 * @docPage hooks/react
 */
import { useEventCallback } from '@fluentui/react-utilities';

/**
 * Alias of Fluent UI 9's `useEventCallback` for a callback with stable identity
 * and an implementation updated by the hook's layout effect.
 *
 * No dependency array, SPFx provider or FluentProvider is required. Use the
 * returned function for event handlers and callbacks invoked after commit,
 * never during rendering. Reactive effect dependencies must remain explicit;
 * this hook is not a universal replacement for React's `useCallback`.
 *
 * @param fn - Callback to invoke, with its current captured props and state
 * after the hook's layout effect runs. Arguments and return types are inferred.
 * @returns A stable callback that forwards arguments, return values (including
 * promises) and thrown errors unchanged. The export is the Fluent hook itself,
 * with no wrapper or additional behavior.
 * @example Stable increment handler
 * ```tsx
 * import * as React from 'react';
 * import { useStableCallback } from '@apvee/spfx-react-toolkit';
 *
 * function Counter() {
 *   const [count, setCount] = React.useState(0);
 *   const increment = useStableCallback(() => setCount(count + 1));
 *   return <button onClick={increment}>{count}</button>;
 * }
 * ```
 * @see https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-utilities/src/hooks/useEventCallback.ts
 * @useWhen An event handler needs stable identity and the latest committed implementation.
 * @avoidWhen A callback must run during rendering, or an effect must re-run when captured data changes; identity stays stable and the implementation updates only after commit.
 * @typeParam Args - Tuple of callback argument types inferred from fn.
 * @typeParam Return - Callback return type, including a promise when fn is async.
 * @remarks
 * Before the first isomorphic layout effect, invocation throws; during a later render it can still call the previous implementation. A state update does not immediately change the implementation before commit. An async call keeps its starting closure across await; later commits affect later invocations only. No loading, rejection handling or cancellation is added. The shared `@fluentui/react-utilities` peer supplies the hook and needs no provider.
 * @example Report current count through a retained callback
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 *
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [count, setCount] = React.useState(0);
 *   const [observed, setObserved] = React.useState<number>();
 *   const [message, setMessage] = React.useState('');
 *   const [sameIdentity, setSameIdentity] = React.useState<boolean>();
 *   const report = useStableCallback((label: string) => {
 *     setObserved(count);
 *     return `${label}: ${count}`;
 *   });
 *   const originalReport = React.useRef<typeof report>();
 *   React.useEffect(() => {
 *     if (originalReport.current === undefined) originalReport.current = report;
 *   }, [report]);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setCount(0);
 *     setObserved(undefined);
 *     setMessage('');
 *     setSameIdentity(undefined);
 *   }, [resetRevision]);
 *   const invokeOriginal = (): void => {
 *     const original = originalReport.current;
 *     if (original === undefined) return;
 *     setSameIdentity(original === report);
 *     setMessage(original('Current count'));
 *   };
 *   return render({
 *     commands: <>
 *       <Button onClick={() => setCount(current => current + 1)}>Increment counter</Button>
 *       <Button onClick={() => setMessage(report('Current count'))}>Report</Button>
 *       <Button onClick={invokeOriginal}>Invoke original callback</Button>
 *     </>,
 *     status: <div role="status" aria-live="polite">
 *       <p>Current count: {count}</p>
 *       <p>Observed counter: {observed === undefined ? 'not invoked' : observed}</p>
 *       <p>{message}</p>
 *       <p>Same callback identity: {sameIdentity === undefined ? 'not checked' : sameIdentity ? 'yes' : 'no'}</p>
 *     </div>,
 *     diagnostics: <p>Invoke after commit. Reset keeps this mounted component and its original callback reference.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Increment twice, then invoke the original callback: it reports 2. Increment again; the observed value changes only when invoked.',
 * };
 * ```
 * @example Retain callback identity through rerenders
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
 *
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 *   children?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 *
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const [count, setCount] = React.useState(0);
 *   const [observed, setObserved] = React.useState<number>();
 *   const [message, setMessage] = React.useState('');
 *   const [sameIdentity, setSameIdentity] = React.useState<boolean>();
 *   const report = useStableCallback((label: string) => {
 *     setObserved(count);
 *     return `${label}: ${count}`;
 *   });
 *   const originalReport = React.useRef<typeof report>();
 *   React.useEffect(() => {
 *     if (originalReport.current === undefined) originalReport.current = report;
 *   }, [report]);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setCount(0);
 *     setObserved(undefined);
 *     setMessage('');
 *     setSameIdentity(undefined);
 *   }, [resetRevision]);
 *   const invokeOriginal = (): void => {
 *     const original = originalReport.current;
 *     if (original === undefined) return;
 *     setSameIdentity(original === report);
 *     setMessage(original('Current count'));
 *   };
 *   return render({
 *     commands: <>
 *       <Button onClick={() => setCount(current => current + 1)}>Increment counter</Button>
 *       <Button onClick={() => setMessage(report('Current count'))}>Report</Button>
 *       <Button onClick={invokeOriginal}>Invoke original callback</Button>
 *     </>,
 *     status: <div role="status" aria-live="polite">
 *       <p>Current count: {count}</p>
 *       <p>Observed counter: {observed === undefined ? 'not invoked' : observed}</p>
 *       <p>{message}</p>
 *       <p>Same callback identity: {sameIdentity === undefined ? 'not checked' : sameIdentity ? 'yes' : 'no'}</p>
 *     </div>,
 *     diagnostics: <p>Invoke after commit. Reset keeps this mounted component and its original callback reference.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Increment and invoke the original callback after rerenders. Identity stays yes, including after Reset clears the values without remounting.',
 * };
 * ```

 */
export const useStableCallback = useEventCallback;
