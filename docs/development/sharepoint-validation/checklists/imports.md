# Imports

[Validation index](../INDEX.md). These procedures require an authenticated host; their presence is not an executed result.


All authenticated rows are **NOT EXECUTED** until a tenant run records them. Build library before app and restart an active serve, then open the lazy **Imports** panel. Keep the existing React Hooks and Styles scenarios in the regression run. See [package imports](../../../guides/package-imports.md) for the compatibility contract.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Root/domain/legacy callbacks | In each card increment twice, then invoke the original captured callback. Observed counter is 2 and identity is yes; the other counters remain independent. Increment and invoke again to confirm current committed state. | NOT EXECUTED |
| Mixed descriptor composition | Compare all three preview widths and body typography. Apply override in each: 360px replaces 240px. Remove it: 240px returns. Mixed root/facade/leaf descriptors emit equivalent active classes/computed styles. | NOT EXECUTED |
| Shared provider direction | Switch right-to-left and back. All previews use the same provider direction; logical start padding moves from left to right and back while width and typography remain equivalent. Record actual computed styles and provider context behavior. | NOT EXECUTED |
| Lifecycle and isolation | Open/close/reopen Imports, and exercise it in two web parts. Counters and provider direction remain scoped to the mounted sample instance. Check console warnings/errors and existing root panels. CSS remaining in the renderer after unmount is allowed by the existing cache/lifetime contract. | NOT EXECUTED |

Record local Node interaction tests, any actual standalone browser fixture execution, and the authenticated SPFx sample separately. The broad sample bundle cannot prove minimal-consumer tree shaking. Direct PnP callers must import their own features; any tenant operations still require actual authentication and effective permissions. No registration, packaging or local probe result establishes those grants.
