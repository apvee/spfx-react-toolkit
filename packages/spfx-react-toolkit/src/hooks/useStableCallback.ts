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
 * @example
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
 */
export const useStableCallback = useEventCallback;
