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
 * import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
 *
 * export function Counter(): React.ReactElement {
 *   const [count, setCount] = React.useState(0);
 *   const [message, setMessage] = React.useState('');
 *   const report = useStableCallback((label: string) => `${label}: ${count}`);
 *   const originalReport = React.useRef(report);
 *   return <div>
 *     <button onClick={() => setCount(current => current + 1)}>Increment</button>
 *     <button onClick={() => setMessage(report('Current count'))}>Report</button>
 *     <button onClick={() => setMessage(originalReport.current('Current count'))}>Invoke original callback</button>
 *     <p>Count: {count}</p>
 *     <p>{message}</p>
 *     <p>Same callback identity: {originalReport.current === report ? 'yes' : 'no'}</p>
 *   </div>;
 * }
 * ```
 */
export const useStableCallback = useEventCallback;
