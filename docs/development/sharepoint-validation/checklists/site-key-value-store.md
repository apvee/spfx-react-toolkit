# Site collection key-value store

[Validation index](../INDEX.md). These procedures require an authenticated host; their presence is not an executed result.


All rows below are **NOT EXECUTED** until an authenticated run records evidence. Use the Site panel's manual actions and a unique disposable `toolkit-site-demo-<run>-` prefix. Record the collection root and current web URLs with sensitive details redacted, account roles, effective root/list grants, Network requests/statuses, UI outcomes and cleanup. An instrumented component or standalone service can exercise cases the panel cannot configure. See the [hook](../../../api/hooks/storage/site-key-value-store.md#usespfxsitekeyvaluestore) and [service](../../../api/services/site-key-value-store.md#createspfxsitekeyvaluestoreservice) contracts.

| Check | Procedure and expected observation | Status |
|-------|------------------------------------|--------|
| Root and two subsites | Save a disposable key from one subsite, then get/list/update from root and a second subsite. Network targets only the collection root `SiteKeyValueStore`; all three see the same key. | NOT EXECUTED |
| Collection isolation | Use the same key in two test collections with different values. Each reads its own root store. | NOT EXECUTED |
| Absent store without setup | On an authorized fresh collection, mount, get, list and remove. Capture zero POST requests; get returns undefined, list returns empty and remove is a no-op after confirmed absence. Denied access must show an error rather than successful absence. | NOT EXECUTED |
| Contributor on existing list | With valid schema and effective Add/Edit/Delete Items but no Manage Lists, save/update/remove disposable items. `canWrite` is advisory; confirm actual server success. | NOT EXECUTED |
| Root/list ACL differences | Compare a subsite contributor without root/list grants, a root contributor denied by unique list permissions, and a list contributor. Record indicators and server results. On absent list, root Manage Lists plus item grants is required; schema repair may require additional rights. | NOT EXECUTED |
| Parallel first save | Start first saves concurrently from separate service instances/pages. Confirm one valid hidden list, correct fields and unique keys; record bounded recovery and no duplicate key. | NOT EXECUTED |
| Schema repair and incompatibility | In a dedicated disposable collection/list, test missing fields or indexing/uniqueness repair through save/ensureListReady, wrong field types and duplicate existing keys. Reads never repair; incompatible writes reject without deleting data. | NOT EXECUTED |
| Complete pagination | Seed more than one server page of disposable keys using authorized tooling. List returns every page in server Title order. Deny/fail a later page: service rejects; hook shows error with empty fallback, never partial success. | NOT EXECUTED |
| Errors, retry and overlap | Force read/write failures, restore access and retry. Writes reject without success messages or unhandled rejections; reads show error rather than absence. Overlap operations, change collection/client or unmount; pending flags and latest-started errors follow current identity while dispatched remote writes may finish. | NOT EXECUTED |
| Key/value/description semantics | Use case variants of a disposable key and confirm one stored Title casing. Read blank Note Value:null as empty string and stored text null as null. Check legacy primitive/date/bigint coercions, omitted-description retention and explicit empty-description clearing. | NOT EXECUTED |
| Tenant regression | Repeat tenant catalog CRUD/admin checks above; tenant APIs retain catalog targeting, existing serialization and permission behavior. Site collection grants do not grant tenant catalog writes. | NOT EXECUTED |

Do not delete a pre-existing store or unrelated data to create a test condition. Only use authorized disposable test collections for list/schema setup, and remove keys/items created by this run from each collection. Record blocked or unavailable cases explicitly.
