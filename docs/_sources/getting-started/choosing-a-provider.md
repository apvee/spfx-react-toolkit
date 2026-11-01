# Choose a provider

Choose a provider for the actual initialized SPFx instance, and bind its React tree to that host surface’s lifetime.

[Documentation home](../INDEX.md) · [Full provider contracts](../../api/core/providers.md)

Use the provider matching the actual SPFx instance:

| Host | Provider |
| --- | --- |
| Web part | `SPFxWebPartProvider` |
| Application Customizer | `SPFxApplicationCustomizerProvider` |
| Field Customizer | `SPFxFieldCustomizerProvider` |
| ListView Command Set | `SPFxListViewCommandSetProvider` |

Pass that instance through `instance` and place consumers below the provider. SPFx hooks need provider context; the provider maintains per-instance state and synchronizes supported host lifecycle updates. See the provider reference for each host's mounting and cleanup pattern, and [core types](../../api/core/types.md) for the accepted instance unions and props.

Generic `useStableCallback` and `useSx` need no SPFx provider. Typed styles still need the shared renderer and theme variables when using theme tokens. Helpers receive plain inputs; services receive initialized clients directly. See [non-React composition](../guides/helpers-and-services.md).

The private sample has actual WebPart and Application Customizer entry points. Field Customizer and Command Set coverage currently checks exports; live mounting requires their matching hosts. An exported provider is not evidence that every host has been exercised. See [host validation](../../development/sharepoint-validation/INDEX.md).

## Web part shell composition

For a web part, [WebPartShell](../../api/controls/web-part-shell.md) composes the matching SPFx provider and shared Fluent theme wrapper with an editable title and optional Card presentation. Supply the actual web part instance and its persisted property contract; do not nest a second SPFx provider merely to use hooks in Shell children. Other SPFx host kinds continue to use their matching direct providers. Generic controls document their own Fluent/localization requirements.

## Late Application Customizer placeholders

Placeholders can appear after initialization and be disposed during navigation. Subscribe to `changedEvent`, retry creation when it changes, and bind disposal to the placeholder captured by each callback. A late disposal callback from an old placeholder must not unmount a replacement.

## Provider lifetime and synchronized state

### Shared provider contract

`SPFxWebPartProviderProps`, `SPFxApplicationCustomizerProviderProps`, `SPFxFieldCustomizerProviderProps` and `SPFxListViewCommandSetProviderProps` each require the initialized matching native host `instance`, as the complete provider recipes below demonstrate and accept optional React `children`. The `TProps extends {}` generic defaults to `{}` and describes the host property bag. Omitted children render no application content. Each provider returns a JSX element and waits for its active native ServiceScope to finish before mounting descendants; it does not register, finish or authorize services itself. The internal fallback proceeds if a scope is missing, but hooks requiring native services may then fail.

The exported providers are typed wrappers over the shared runtime; `SPFxProviderBase`, runtime stores and React context objects remain private. Hooks consuming SPFx context/runtime state throw outside the matching provider. The generic React hooks `useStableCallback` and `useSx` do not require an SPFx provider.

### Provider Features

All providers use the common runtime. The sample mounts WebPart and Application Customizer providers in their real host entry points; Field Customizer and ListView Command Set exports are checked but require separate real-host validation. See [SharePoint validation](../../development/sharepoint-validation/INDEX.md).

### Instance Isolation

Each SPFx instance gets its own isolated state store. Provider runtime state is scoped to each SPFx instance. Explicitly shared external clients, tenant data and application objects remain shared by their own contracts.

- **Property Pane → React**: A host render reconciles top-level values, removals and changed references using a shallow snapshot; nested in-place mutations are not deep-observed
- **React → SPFx**: Property updates via `useSPFxProperties` sync back to SPFx

### Theme Subscription

Theme changes (light/dark mode) are automatically detected and propagated to hooks.

### Display Mode Tracking

Edit/Read mode changes are tracked and available via `useSPFxDisplayMode`.

### Lifecycle

The host must unmount React when it disposes its rendered surface. The provider removes its theme subscription and runtime listeners on unmount; scope replacement waits for the current ServiceScope and ignores obsolete readiness callbacks. Local lifecycle regression tests cover these paths, but they do not establish a universal memory or tenant-host guarantee.

### Container Observation

WebPart providers expose the native `domElement`; other current host kinds expose no container element. `useSPFxContainerInfo` subscribes to the available element with ResizeObserver (or the existing browser fallback) and `useSPFxContainerSize` derives breakpoints. Mount the container hook when measurements are needed; the provider itself does not observe an element.

<!-- toolkit:include symbol="SPFxWebPartProvider" section="example" title="Mount and dispose a web part" -->

<!-- toolkit:include symbol="SPFxApplicationCustomizerProvider" section="example" title="Retry late placeholders and dispose captured roots" -->

<!-- toolkit:include symbol="SPFxFieldCustomizerProvider" section="example" title="Render and dispose each field cell" -->

<!-- toolkit:include symbol="SPFxListViewCommandSetProvider" section="example" title="Dispose a temporary command dialog root" -->
