# Imports

[Validation index](../INDEX.md) · [Setup](../setup.md)

All authenticated rows are **NOT EXECUTED** until a tenant run records them. Build library before app and restart an active serve, then open the lazy **Imports** panel. Keep the existing React Hooks and Styles scenarios in the regression run. See [package imports](../../../guides/package-imports.md) for the compatibility contract.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Root/domain/legacy callbacks | Select each adapter in turn, increment twice, then invoke its original captured callback. Observed counter is 2 and identity is yes. Increment and invoke again to confirm current committed state. Switching adapters mounts a new probe; use the legacy comparison fixture for simultaneous independent counters. | NOT EXECUTED |
| Mixed descriptor composition | Compare each selected adapter's preview width and body typography. Apply override: 360px replaces 240px. Remove it: 240px returns. Mixed root/facade/leaf descriptors emit equivalent active classes/computed styles. | NOT EXECUTED |
| Shared provider direction | Switch the selected adapter right-to-left and back beneath the shared provider. Logical start padding moves from left to right and back while width and typography remain equivalent. Repeat for each adapter and record actual computed styles and provider context behavior. | NOT EXECUTED |
| Lifecycle and isolation | Open/close/reopen Imports, and exercise it in two web parts. Counters and provider direction remain scoped to the mounted sample instance. Check console warnings/errors and existing root panels. CSS remaining in the renderer after unmount is allowed by the existing cache/lifetime contract. | NOT EXECUTED |

Record local Node interaction tests, any actual standalone browser fixture execution, and the authenticated SPFx sample separately. The broad sample bundle cannot prove minimal-consumer tree shaking. Direct PnP callers must import their own features; any tenant operations still require actual authentication and effective permissions. No registration, packaging or local probe result establishes those grants.
