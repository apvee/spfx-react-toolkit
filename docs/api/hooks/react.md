# React Utility Hooks

These hooks can be used in React components without an SPFx provider. `useStableCallback` needs no FluentProvider; token-based `useSx` selections need theme CSS variables in scope.

## useStableCallback

`useStableCallback` is a direct alias of Fluent UI 9's `useEventCallback` from `@fluentui/react-utilities`. It adds documentation and a toolkit export, with no wrapper or additional behavior.

### Signature

```ts
const useStableCallback: <Args extends unknown[], Return>(
  fn: (...args: Args) => Return
) => (...args: Args) => Return;
```

- `fn`: the callback to execute. Arguments and return values are inferred, including Promise return types.
- Result: a callback with stable identity across renders of the mounted component. It forwards arguments, results and thrown errors unchanged.

### Example

```tsx
import * as React from 'react';
import { useStableCallback } from '@apvee/spfx-react-toolkit';

export function Counter() {
  const [count, setCount] = React.useState(0);
  const report = useStableCallback((label: string) => `${label}: ${count}`);
  const [message, setMessage] = React.useState('');

  return (
    <div>
      <button onClick={() => setCount(current => current + 1)}>Increment</button>
      <button onClick={() => setMessage(report('Current count'))}>Report</button>
      <p>{count}</p>
      <p>{message}</p>
    </div>
  );
}
```

### Timing and limitations

Fluent updates the internal callback in an isomorphic layout effect. Browser event handlers invoked after that effect read the updated captured props and state. Updating state does not immediately update the callback before React renders and runs that effect.

Do not invoke the returned callback during rendering. Before the first layout effect it throws; during later renders it may still call the previous implementation. Use it for event handlers and external callbacks invoked after commit.

The stable identity does not notify effects that a captured value changed. Keep reactive effect dependencies explicit: an effect that fetches data for `listId` must still depend on `listId`. This hook is not a universal replacement for `useCallback`.

An async invocation retains the closure with which it started, including across `await`; later renders affect subsequent invocations, not the invocation already running. No loading state, error handling or cancellation is added.

### Dependencies

The toolkit declares `@fluentui/react-utilities@^9.25.1` as a mandatory shared peer, alongside its existing Fluent theme peers. Install compatible versions in the consuming host. No provider is required by this hook.

The underlying hook is exported and typed by Fluent UI 9, but its upstream source includes an `@internal` JSDoc annotation. The toolkit preserves the upstream behavior rather than maintaining a separate implementation.

### Web part scenario

Open **React Hooks** in the test web part. Increment the counter, then select **Invoke original callback**. The observed counter should match the current counter and **Same callback identity** should read **yes**. The scenario retains the callback captured at mount rather than replacing it on each render.

Local tests exercise that component in a DOM harness with the actual hook. Authenticated SharePoint-host execution must be recorded separately.

### Source

- [Toolkit alias](../../../packages/spfx-react-toolkit/src/hooks/useStableCallback.ts)
- [Fluent implementation](https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-utilities/src/hooks/useEventCallback.ts)

## useSx

```ts
function useSx(options?: SxOptions): SxFunction;
```

`options` optionally declares the effective `dir: 'ltr' | 'rtl'`. Otherwise the hook uses Fluent direction context, with LTR as the context-free default. The returned `SxFunction` composes immutable descriptors, existing class strings and conditional falsy inputs into a className string. It is stable while renderer and effective direction stay unchanged. Empty composition returns `''`.

```tsx
import { useSx, flex, gap, presets, padding, foreground } from '@apvee/spfx-react-toolkit';

export function Surface() {
  const sx = useSx();
  return <div className={sx(flex.row, gap.medium, presets.canvas, padding.medium, foreground.subtle)}>Content</div>;
}
```

Use it at the top level of a component. The returned composer is an ordinary function usable multiple times during rendering. No SPFx provider or additional renderer is created. Theme-based selections need existing Fluent CSS variables, usually from the host's FluentProvider.

See the [complete styles reference](../helpers/styles.md) for all qualified exports, mapping, numeric domains, state/query priority, external-class behavior, direction, container ancestor requirements, theme restrictions and cache growth. Shared peers are `@griffel/core@^1.19.2`, `@griffel/react@^1.5.30`, `@fluentui/react-shared-contexts@^9.25.2`, alongside the existing Fluent peers.

The lazy **Styles** web part panel exercises the catalog, two independent query regions, viewport scopes, interaction states, themes and native scrolling. Local React/Griffel tests and browser observations are separate from authenticated SharePoint-host execution, which is **NOT EXECUTED** until recorded in the [host checklist](../../SHAREPOINT-VALIDATION.md#styles-and-usesx).
