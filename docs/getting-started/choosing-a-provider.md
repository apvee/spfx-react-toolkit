# Choose a provider

[Documentation home](../INDEX.md) · [Full provider contracts](../api/core/providers.md)

Use the provider matching the actual SPFx instance:

| Host | Provider |
| --- | --- |
| Web part | `SPFxWebPartProvider` |
| Application Customizer | `SPFxApplicationCustomizerProvider` |
| Field Customizer | `SPFxFieldCustomizerProvider` |
| ListView Command Set | `SPFxListViewCommandSetProvider` |

Pass that instance through `instance` and place consumers below the provider. SPFx hooks need provider context; the provider maintains per-instance state and synchronizes supported host lifecycle updates. See the provider reference for each host's mounting and cleanup pattern, and [core types](../api/core/types.md) for the accepted instance unions and props.

Generic `useStableCallback` and `useSx` need no SPFx provider. Typed styles still need the shared renderer and theme variables when using theme tokens. Helpers receive plain inputs; services receive initialized clients directly. See [non-React composition](../guides/helpers-and-services.md).

The private sample has actual WebPart and Application Customizer entry points. Field Customizer and Command Set coverage currently checks exports; live mounting requires their matching hosts. An exported provider is not evidence that every host has been exercised. See [host validation](../development/sharepoint-validation/INDEX.md).

## Web part shell composition

For a web part, [WebPartShell](../api/controls/web-part-shell.md) composes the matching SPFx provider and shared Fluent theme wrapper with an editable title and optional Card presentation. Supply the actual web part instance and its persisted property contract; do not nest a second SPFx provider merely to use hooks in Shell children. Other SPFx host kinds continue to use their matching direct providers. Generic controls document their own Fluent/localization requirements.
