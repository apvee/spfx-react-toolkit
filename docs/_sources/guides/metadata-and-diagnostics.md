# Read host metadata and diagnose behavior

Use provider hooks for live host integration and explicit helpers for supplied PageContext snapshots. Distinguish UI culture, regional formatting, site correlation, measurement and asynchronous SDK metadata.

[Documentation home](../INDEX.md)

## Typed host boundaries

Pass the actual initialized host instance to its matching provider. `HostKind`, `SPFxComponent`, `SPFxContextType` and shared provider/context types describe those boundaries; they do not make private runtime stores or React contexts public. Use the generated [core reference](../../api/core/types.md) for their shapes and legal type imports. A supplied generic property model is a compile-time expectation, not validation of a host property bag.

## Host context and instance identity

Use {@link useSPFxContext} for the current native context and {@link useSPFxInstanceInfo} for the actual host/instance identity. Shared tenant data and explicitly shared clients do not become isolated merely because a provider's React runtime store is per instance. Do not use hook return objects as a universal subscription to mutable nested PageContext fields.

<!-- toolkit:include symbol="useSPFxContext" section="example" title="Render host metadata and guarded host-specific content" -->

## Consume a registered service

The host owns registration and finishing its ServiceScope. Mount the matching provider after host initialization; it waits for scope readiness before mounting consumers. Native `consume` semantics still apply, including service defaults and failures. Consume stable dependencies inside memoized React work where appropriate; do not finish/register services from render.

<!-- toolkit:include symbol="useSPFxServiceScope" section="example" title="Consume a registered service after native scope readiness" -->

## Environment and asynchronous SDK context

The synchronous page-context helper classifies supplied metadata; the provider hook can acquire the native SDK asynchronously. Local, Teams, Office and Outlook indicators can overlap, and the selected category has precedence. `useSPFxTeams` has its historical name but can serve Teams, Office and Outlook. Narrow its unknown context before using host-specific fields.

An acquired SDK context does not guarantee live theme delivery: the SDK exposes a single registration handler without a public unregister API, and another consumer can replace it. Listener cleanup and ignored obsolete completions do not promise that every external SDK callback disappears. Use {@link useSPFxEnvironmentInfo} to identify the actual host and canonical hook docs for update/lifetime limits.

## Culture and time-zone choices

A UI culture controls native text, a regional locale controls formatting, and a verified numeric SharePoint group can provide a representative IANA zone. Locale `it-IT` does not imply Rome; group 4 maps to `Europe/Berlin`. Passing undefined to Intl or RelativeTime deliberately uses the runtime zone. Choose an explicit unavailable-zone label when that fallback would mislead readers.

The [audited mapping and immutable provenance](localization.md#audited-time-zone-mapping) retains all 111 candidates, 104 verified groups, excluded IDs, pinned source revisions, snapshot dates and license notices. Standard groups cannot verify custom SharePoint Server `TIMEZONE.XML` rules.

<!-- toolkit:include symbol="getSPFxLocaleInfo" section="example" title="Choose an explicit missing site-zone policy" -->

## Theme subscriptions and container measurement

The native theme subscription and Fluent theme projection have separate roles. Treat shared returned theme objects as immutable, and establish Fluent theme variables through the host provider before using token descriptors. Only web-part hosts currently expose their native container; measuring hooks subscribe when mounted. A numeric width helper does not measure, and CSS container scopes have inclusive thresholds that differ from the helper's exclusive classification thresholds.

<!-- toolkit:include symbol="getSPFxContainerSize" section="example" -->

## User and site ownership

User metadata is not a profile request; optional email/classification/logo/group/list fields depend on host data. User-photo hooks own their created object URLs and revoke them on replacement/unmount. Standalone photo services return caller-owned Blobs instead. Do not retain a hook URL beyond its identity/lifetime or infer authentication from the presence of a user name.

Site collection correlation is distinct from an AAD tenant-origin assumption. The combined helper recipe reads supplied user/site/list/locale/environment/page/correlation metadata without acquiring SDK context, and leaves each authority explicit.

<!-- toolkit:include symbol="getSPFxSiteInfo" section="example" title="Read combined host metadata" -->

## Property editing and host reconciliation

Partial `setProperties` updates differ from an updater returning the complete next property bag. Use a previous-state updater for counters/timestamps and use edit/read metadata to expose configuration deliberately. Host Property Pane reconciliation compares top-level values/references; nested in-place mutation is not a deep-observed edit. Hooks update the SPFx bag, while publication and server persistence retain their host contracts.

<!-- toolkit:include symbol="useSPFxProperties" section="example" title="Update selected properties" -->

<!-- toolkit:include symbol="useSPFxProperties" section="example" title="Update properties from the previous snapshot" -->

<!-- toolkit:include symbol="useSPFxProperties" section="example" title="Edit configuration in native edit mode" -->

## Performance and logging

Use {@link useSPFxPerformance} to time an actual operation and await its callback. Concurrent same-name timing calls use distinct internal marks; explicit public mark names and browser entries remain caller-owned. A zero-duration fallback is not evidence that work was free. Callback failures still reject. {@link useSPFxLogger} adds instance/correlation context to native logging without creating a remote telemetry transport; avoid secrets and tokens in diagnostic messages.

<!-- toolkit:include symbol="useSPFxPerformance" section="example" -->

Local lifecycle, metadata and timing checks do not execute every native host. Use the [host procedures](../development/sharepoint-validation/INDEX.md) and record real outcomes separately.
