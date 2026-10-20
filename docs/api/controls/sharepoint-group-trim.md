# SharePointGroupTrim

`SharePointGroupTrim` is a headless current-user access trim. Import the compound value and its types from the package root or controls barrel. It requires a matching SPFx provider; place it within the provider appropriate for your actual host.

```tsx
import * as React from 'react';
import { SharePointGroupTrim } from '@apvee/spfx-react-toolkit';

function ProtectedAction() {
  return <SharePointGroupTrim groups={[{ name: 'Project Readers' }]}>
    <SharePointGroupTrim.Loading><p>Checking access</p></SharePointGroupTrim.Loading>
    <SharePointGroupTrim.Fallback><p>Access not observed</p></SharePointGroupTrim.Fallback>
    <SharePointGroupTrim.ErrorFallback><p>Access could not be checked</p></SharePointGroupTrim.ErrorFallback>
    <button type="button">Protected action</button>
  </SharePointGroupTrim>;
}
```

## Direct membership contract

| Qualified member | Contract |
| --- | --- |
| `SharePointGroupTrimProps.groups` | Required nonempty readonly `SharePointGroupSelector[]`; every reference must resolve. |
| `SharePointGroupSelector.id` | Positive Int32 site-group ID, mutually exclusive with name. |
| `SharePointGroupSelector.name` | Exact nonblank group title, mutually exclusive with id. The other selector member is `never`. |
| `SharePointGroupTrimProps.webUrl` | Optional absolute HTTP(S) web URL without credentials/query/fragment; omitted uses current web. |
| `SharePointGroupTrimProps.match` | Optional `all` (default) or `any`; aggregation happens after all requested reads succeed. |
| `SharePointGroupTrimProps.children` | Optional protected content and direct branches; omission renders empty. |
| `SharePointGroupTrimBranchProps.children` | Optional active branch content; omission renders empty. |

Requires a matching SPFx provider and native SPHttpClient. Reads the selected web's current-user site-local ID and direct users for every selected site group across all pages. This does not expand Entra principals nested in SharePoint groups and does not query a transitive directory membership. Exact names use escaped Title filters, local exact equality and all-page ambiguity checks. Missing or ambiguous groups, invalid IDs, HTTP/payload/paging failures are errors; an earlier positive membership cannot hide a later failure. Remote-web IDs belong to that web's site collection.

## Structural branches and lifetime

Place `Loading`, `Fallback` and `ErrorFallback` once each as **direct** children of their own family. They are structural content markers, not Fluent slots. The trim emits caller content without a DOM wrapper. Loading means a current read is pending; Fallback means all requested reads succeeded and access did not match; ErrorFallback means validation, transport, payload or resolution failed. Errors never select Fallback. Every omitted branch renders empty, including an omitted ErrorFallback. An absent `children` prop renders empty.

Inactive protected/branch components never render, run effects or mount. Caller DOM, class names, refs and formatting belong on the caller's children. Controls and markers expose no public DOM, class/style/ref or callback props. Changing the active role or the structural marker's React key remounts its content; the same role/key preserves caller state and child keys. Child arrays/keys remain meaningful. React 17's broad `ReactNode` type can admit functions through its historical `{}` alternative; callbacks are unsupported and never executed.

Duplicate, wrong-family and detectable fragment-wrapped markers fail closed. Keep markers direct rather than hiding them inside wrappers/fragments or custom components. Markers are supported only as direct children of their own trim family. A standalone marker renders its children unconditionally and performs no access check; standalone usage is unsupported. Custom components are ordinary protected children, not a way to declare a structural branch.

Provider, native client, PageContext/current-user identity and target changes hide previous asynchronous grants during render. Obsolete/unmounted requests cannot restore them. Identical inflight reads may be shared within the matching provider only; settled grants are not cached and there is no polling. Host changes still require a React render. Run/repeat in the sample deliberately changes the trim key to perform a fresh check. UI trimming is advisory; each server operation still needs authorization.

The lazy **SharePoint Group Trim** sample uses this actual control and native APIs, editable selectors, enum fields, all three branches, a protected counter and effect mount gauge. Local fixtures replace only unavailable transports and do not establish authenticated tenant behavior. See [development verification](../../DEVELOPMENT.md) and the [controls index](./INDEX.md).

`SharePointGroupTrimBranchProps` describes each structural marker. `SharePointGroupTrimProps` describes the control.
