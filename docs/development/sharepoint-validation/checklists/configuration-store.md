# Configuration store

[Validation index](../INDEX.md) · [Setup](../setup.md)

These checks cover the additive `useSPFxConfigurationStore` and `createSPFxConfigurationStoreService`. The authenticated run on **2026-10-21** used the **Configuration** panel in the demo collection and Site-B after rebuilding the library/app and restarting serve. The earlier tenant/site-store results remain separate. For subsequent runs, use unique `spfx-toolkit-config-demo-<run>-` keys, an explicit ledger of each physical destination and existing authorized test collections/accounts only.

| Check | Procedure and expected observation | Status |
| --- | --- | --- |
| Three modes and explicit targets | Exercise tenant/site/layered reads and permitted mutation targets. Single-store mode disables the other physical scope in the panel. | PASS — 2026-10-21; tenant-only read 20 despite A's site override; site-only read B's 70; layered reads and scoped Save/Remove succeeded |
| Layered resolution | Save tenant 20, read it from both test collections, save site 50 in collection A. A reads 50/source site; B reads 20/source tenant. Then save B's independent site value 70; A remains 50 and B reads 70. | PASS — 2026-10-21; values were numbers and provenance matched each physical destination |
| Lazy discovery and no implicit requests | Capture no catalog request for a layered site hit or site-only operation, and no requests on mount/selection changes. | NOT EXECUTED on host — network trace not captured; covered by local controlled tests |
| Removal and inheritance | Remove A's site override and read tenant 20. Keep B's override 70, remove tenant: A is absent while B still reads 70/source site. Only selected records are deleted. | PASS — 2026-10-21; A's physical site absence and tenant fallback were checked separately |
| Present values and metadata | Check false, zero, empty text and JSON objects as complete site overrides, without property merge. Key casing/description/ID/source all describe the winning row. Empty description clears destination metadata. | NOT EXECUTED |
| Inventories and errors | Physical inventory shows only selected scope, count and provenance. With an already available restricted account, verify access failures are errors, not successful absence or fallback. No new grants/ACL changes are authorized by this checklist. | NOT EXECUTED |
| Context and conflicts | Change available collection/client context during requests; obsolete outcomes do not publish current state. Record observed 412 behavior only if an authorized reproducible race is available; local controlled tests remain separate. | NOT EXECUTED |
| Cleanup | Remove only run-created tenant entries and each collection's site entries; fresh reads/inventories of each physical scope confirm absence. Preserve lists, unrelated keys and the original page/workbench state. | PASS — 2026-10-21; one tenant and two site entries removed; fresh physical reads and final layered reads confirm absence |

The run used only `spfx-toolkit-config-demo-20261021-1500-7e3b`, initially confirmed absent. Screenshots and the destination ledger are retained under `.docs/maintenance/configuration-store-implementation/`. Temporary workbench tabs were closed without saving; the original tab was preserved. No ACL, tenant grant or list deletion was performed. Network/console traces, restricted-account failures, full value/metadata cases and conflicts were not collected or exercised in this run; local tests do not constitute their host acceptance.

Root/subsite sharing remains NOT EXECUTED: these test collections have no existing subsites. Report tests, browser/console/network evidence and authenticated results separately. See [configuration API](../../../api/hooks/storage/configuration-store.md).
