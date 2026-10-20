# SharePointPermissionTrim

`SharePointPermissionTrim` is a headless current-user access trim. Import the compound value and its types from the package root or controls barrel. It requires a matching SPFx provider; place it within the provider appropriate for your actual host.

```tsx
import * as React from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { SharePointPermissionTrim } from '@apvee/spfx-react-toolkit';

function ProtectedAction() {
  return <SharePointPermissionTrim scope="list" listTitle="Tasks" permissions={[SPPermission.addListItems]}>
    <SharePointPermissionTrim.Loading><p>Checking access</p></SharePointPermissionTrim.Loading>
    <SharePointPermissionTrim.Fallback><p>Access not observed</p></SharePointPermissionTrim.Fallback>
    <SharePointPermissionTrim.ErrorFallback><p>Access could not be checked</p></SharePointPermissionTrim.ErrorFallback>
    <button type="button">Protected action</button>
  </SharePointPermissionTrim>;
}
```

## Resource contract

`SharePointPermissionTrimProps` inherits `SPFxScopedPermissionsOptions`. `SharePointPermissionTrimProps.scope` defaults to `web`, and `SharePointPermissionTrimProps.webUrl` defaults to the provider's current web. URLs must be absolute HTTP(S), without credentials/query/fragment. `SharePointPermissionTrimProps.listId`, `SharePointPermissionTrimProps.listTitle` and `SharePointPermissionTrimProps.listPath` are mutually exclusive; exactly one is required for list/item/folder and forbidden for web. IDs are GUIDs with optional paired braces; titles match exact nonblank text; listPath is decoded web-relative root without a leading slash, query, backslash or dot segments. `SharePointPermissionTrimProps.itemId` is a positive integer required only for item; `SharePointPermissionTrimProps.folderPath` is a decoded server-relative folder path required only for folder. A folder must resolve to a list item belonging to the selected list. Use list scope for a library root folder without a list item.

| Qualified member | Contract |
| --- | --- |
| `SharePointPermissionTrimProps.permissions` | Required nonempty readonly native SPPermission constants or combined masks, with public unsigned Int32 value.High/value.Low. Empty, sparse or malformed inputs are errors before requests. Compatible peer masks are accepted by shape. |
| `SharePointPermissionTrimProps.match` | Optional `all` (default) requires every permission; `any` requires one. |
| `SharePointPermissionTrimProps.children` | Optional protected content and direct branches; omission renders empty. |
| `SharePointPermissionTrimBranchProps.children` | Optional active branch content; omission renders empty. |

Available current-web masks use `useSPFxPermissions` without REST and inherit PageContext freshness. A missing mask or remote web uses native SPHttpClient. Scoped resources read their own effective permissions, not unrelated masks. HTTP/payload failures produce ErrorFallback with no partial grants. Requires a matching SPFx provider and its native PageContext/SPHttpClient; no PnP registrations are required.

## Structural branches and lifetime

Place `Loading`, `Fallback` and `ErrorFallback` once each as **direct** children of their own family. They are structural content markers, not Fluent slots. The trim emits caller content without a DOM wrapper. Loading means a current read is pending; Fallback means all requested reads succeeded and access did not match; ErrorFallback means validation, transport, payload or resolution failed. Errors never select Fallback. Every omitted branch renders empty, including an omitted ErrorFallback. An absent `children` prop renders empty.

Inactive protected/branch components never render, run effects or mount. Caller DOM, class names, refs and formatting belong on the caller's children. Controls and markers expose no public DOM, class/style/ref or callback props. Changing the active role or the structural marker's React key remounts its content; the same role/key preserves caller state and child keys. Child arrays/keys remain meaningful. React 17's broad `ReactNode` type can admit functions through its historical `{}` alternative; callbacks are unsupported and never executed.

Duplicate, wrong-family and detectable fragment-wrapped markers fail closed. Keep markers direct rather than hiding them inside wrappers/fragments or custom components. Markers are supported only as direct children of their own trim family. A standalone marker renders its children unconditionally and performs no access check; standalone usage is unsupported. Custom components are ordinary protected children, not a way to declare a structural branch.

Provider, native client, PageContext/current-user identity and target changes hide previous asynchronous grants during render. Obsolete/unmounted requests cannot restore them. Identical inflight reads may be shared within the matching provider only; settled grants are not cached and there is no polling. Host changes still require a React render. Run/repeat in the sample deliberately changes the trim key to perform a fresh check. UI trimming is advisory; each server operation still needs authorization.

The lazy **SharePoint Permission Trim** sample uses this actual control and native APIs, editable selectors, enum fields, all three branches, a protected counter and effect mount gauge. Local fixtures replace only unavailable transports and do not establish authenticated tenant behavior. See [development verification](../../DEVELOPMENT.md) and the [controls index](./INDEX.md).

## AND composition

Nest actual trims to require both decisions. The inner check mounts only when the outer check allows, so an outer denial/error does not start the inner request. The Permission Trim sample includes this selectable composition. Its nested Entra check additionally needs the native Graph client and approved User.Read/GroupMember.Read.All access.

```tsx
import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';
<SharePointPermissionTrim permissions={[SPPermission.viewListItems]}>
  <EntraGroupTrim groups={[{ name: 'Project Readers' }]}>
    <p>Permission AND Entra membership observed</p>
    <EntraGroupTrim.Fallback>Group access not observed</EntraGroupTrim.Fallback>
    <EntraGroupTrim.ErrorFallback>Group check failed</EntraGroupTrim.ErrorFallback>
  </EntraGroupTrim>
  <SharePointPermissionTrim.Fallback>Permission not observed</SharePointPermissionTrim.Fallback>
  <SharePointPermissionTrim.ErrorFallback>Permission check failed</SharePointPermissionTrim.ErrorFallback>
</SharePointPermissionTrim>;
```

`SharePointPermissionTrimBranchProps` describes each structural marker. `SharePointPermissionTrimProps` describes the control.
