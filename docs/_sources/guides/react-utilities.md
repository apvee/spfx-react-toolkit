# Use stable callbacks and committed values

Keep callback identity stable after commit while reading current state, and distinguish render-time behavior from retained event callbacks.

[Documentation home](../INDEX.md)

## Stable callbacks after commit

### Import choices

Root imports, `/hooks` and historical hook leaves remain supported. The narrow clean `/hooks/useStableCallback` and `/styles/useSx` aliases point to the same canonical hooks. `/styles` exposes `useSx` alongside descriptors and types. Mixed aliases share context and renderer identity rather than adding instances. See [package imports](package-imports.md) for complete examples, resolver/peer requirements and measured budgets. The Imports web part complements the generated useStableCallback scenarios and the existing Styles scenarios.

### useStableCallback

`useStableCallback` is a direct alias of Fluent UI 9's `useEventCallback` from `@fluentui/react-utilities`. It adds documentation and a toolkit export, with no wrapper or additional behavior.

- Result: a callback with stable identity across renders of the mounted component. It forwards arguments, results and thrown errors unchanged.

### Timing and limitations

Fluent updates the internal callback in an isomorphic layout effect. Browser event handlers invoked after that effect read the updated captured props and state. Updating state does not immediately update the callback before React renders and runs that effect.

Do not invoke the returned callback during rendering. Before the first layout effect it throws; during later renders it may still call the previous implementation. Use it for event handlers and external callbacks invoked after commit.

The stable identity does not notify effects that a captured value changed. Keep reactive effect dependencies explicit: an effect that fetches data for `listId` must still depend on `listId`. This hook is not a universal replacement for `useCallback`.

An async invocation retains the closure with which it started, including across `await`; later renders affect subsequent invocations, not the invocation already running. No loading state, error handling or cancellation is added.

The toolkit declares `@fluentui/react-utilities@^9.26.7` as a mandatory shared peer, alongside its existing Fluent theme peers. Install compatible versions in the consuming host. No provider is required by this hook.

The underlying hook is exported and typed by Fluent UI 9, but its upstream source includes an `@internal` JSDoc annotation. The toolkit preserves the upstream behavior rather than maintaining a separate implementation.

### Web part scenario

Open **Hook → useStableCallback** in the test web part and select either source-owned scenario. Increment the counter, then select **Invoke original callback**. The observed counter should match the current counter and **Same callback identity** should read **yes**. The scenario retains the callback captured at mount rather than replacing it on each render.

Local tests exercise that component in a DOM harness with the actual hook. Authenticated SharePoint-host execution must be recorded separately.

- [Toolkit alias](../../../packages/spfx-react-toolkit/src/hooks/useStableCallback.ts)
- [Fluent implementation](https://github.com/microsoft/fluentui/blob/master/packages/react-components/react-utilities/src/hooks/useEventCallback.ts)

<!-- toolkit:include symbol="useStableCallback" section="example" title="Report current count through a retained callback" -->

<!-- toolkit:include symbol="useStableCallback" section="example" title="Retain callback identity through rerenders" -->
