# Graph and storage permissions

[Validation index](../INDEX.md). These procedures require an authenticated host; their presence is not an executed result.


The checked-in [package solution config](../../../../apps/spfx-react-toolkit-test/config/package-solution.json) does not declare `webApiPermissionRequests`. A served sample does not grant Graph/custom API permissions. Before positive Graph tests, have the tenant administrator arrange the required delegated permissions through the established SPFx approval workflow for an authorized test solution. Permission/precheck output is evidence of token acquisition for the current runtime, not a tenant grants inventory or server authorization guarantee.

For OneDrive app-root reads/writes, use the applicable delegated `Files.ReadWrite.AppFolder` or `Files.ReadWrite` scope. Photo/client features need their appropriate endpoint permissions; custom API prechecks must use the registered resource and scope. Record exactly which approvals/account/endpoint were used rather than treating one successful token as universal permission.

| Storage check | Procedure and expected observation |
|---------------|------------------------------------|
| Browser defaults | Use a disposable test component with an inline object default. Missing key settles without a render loop; changing default alone preserves value. `remove()` resets to the latest default. Change key and inspect its persisted value/current fallback. |
| Browser negative cases | Block storage or introduce invalid JSON for a disposable scoped key. The hook falls back without throwing; writes are best effort. Exercise a matching storage event in another eligible browser context. |
| OneDrive lazy seed | With `autoFetch: false` and `defaultValue`, initial data is the seed and no content request starts until load. Change file/folder and confirm the new seed. |
| OneDrive missing/current identity | Read a missing disposable JSON file with createIfMissing, confirm only the current missing file is created. Change file during an old read; old 404 must not create the current file or overwrite newer local data. |
| OneDrive failure | Denied Graph scope, missing drive, invalid JSON and network failure expose the relevant read/write error. Write promises reject; read failures are captured in state. |
| Tenant properties | Provision `spfx-toolkit-test-version` and `spfx-toolkit-test-counter` via the tenant's established admin tooling, then load in Tenant panel. Also test a missing key and inaccessible catalog. Properties are read-only through this API. |
| Tenant key-value CRUD | As an authorized catalog Site Collection Admin, save a disposable key, get/list it, update and remove it. First write may provision hidden TenantKeyValueStore. Repeat with a regular user: canWrite is false and unauthorized writes fail. Reads depend on actual catalog/list access. |
| Tenant read failure | Inaccessible/unprovisioned catalog and denied list requests produce errors. `get`/`list` may return undefined/empty fallback; inspect error before calling that an absent record. Loading tracks pending read/write channels separately. |
| Legacy serialization | Read numeric/boolean/null-like strings, objects, dates and large bigint inputs. Document observed coercion/loss; there is no complete type round-trip guarantee and generic types do not validate stored data. |
