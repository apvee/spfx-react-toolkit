# EntraGroupTrim

`EntraGroupTrim` is a headless current-user access trim. Import the compound value and its types from the package root or controls barrel. It requires a matching SPFx provider; place it within the provider appropriate for your actual host.

```tsx
import * as React from 'react';
import { EntraGroupTrim } from '@apvee/spfx-react-toolkit';

function ProtectedAction() {
  return <EntraGroupTrim groups={[{ name: 'Engineering' }]} groupType="M365" relationship="memberOrOwner">
    <EntraGroupTrim.Loading><p>Checking access</p></EntraGroupTrim.Loading>
    <EntraGroupTrim.Fallback><p>Access not observed</p></EntraGroupTrim.Fallback>
    <EntraGroupTrim.ErrorFallback><p>Access could not be checked</p></EntraGroupTrim.ErrorFallback>
    <button type="button">Protected action</button>
  </EntraGroupTrim>;
}
```

## Directory contract

| Qualified member/type | Contract |
| --- | --- |
| `EntraGroupTrimProps.groups` | Required nonempty readonly `EntraGroupSelector[]`; invalid selectors fail before requests. |
| `EntraGroupSelector.id` | Group GUID, compared case-insensitively; mutually exclusive with name. |
| `EntraGroupSelector.name` | Exact nonblank displayName; mutually exclusive with id. The other selector member is `never`. |
| `EntraGroupSelector.groupType` | `never`; classification belongs only on the control. |
| `EntraGroupTrimProps.groupType`, `EntraGroupType` | Optional case-sensitive `All` (default), `M365` or `Security`. M365 requires Unified; Security requires securityEnabled and excludes Unified, including mail-enabled/dynamic security groups. All adds no metadata restriction. |
| `EntraGroupTrimProps.relationship`, `EntraGroupRelationship` | Optional `member` (default), `owner`, or `memberOrOwner`. |
| `EntraGroupTrimProps.match` | Optional `all` (default) or `any`, after all requested reads succeed. |
| `EntraGroupTrimProps.children` | Optional protected content and direct branches; omission renders empty. |
| `EntraGroupTrimBranchProps.children` | Optional active branch content; omission renders empty. |

Requires a matching SPFx provider with its native MSGraphClientV3 factory. All-ID member checks do not resolve group existence: an unreturned GUID is denied and cannot distinguish a missing group from nonmembership. All-ID member checks use `/me/checkMemberGroups` with deduplicated chunks of at most twenty and avoid metadata reads. Graph membership is transitive for security groups; Microsoft 365 group membership is direct. Names query escaped exact displayName, inspect all pages, apply the requested groupType **locally before ambiguity**, and fail on missing/ambiguous matches. A specific typed ID requires metadata; a mismatch is an error, not a denied check. No advanced query headers/eventual-index assumptions are added.

Owners read `/me` and every direct owners page; ID-only limited-information objects suffice. Metadata-confirmed synchronized groups, distribution groups and non-Unified mail-enabled groups have unsupported ownership and select ErrorFallback before member/owner reads, including memberOrOwner. Every requested read must succeed; `any` or a positive member result cannot hide another read failure. HTTP/Graph/payload errors never become partial grants. Hidden membership may be omitted by Graph: absence in a response is not proof of real-world absence.

The sample solution config requests Microsoft Graph `User.Read` and `GroupMember.Read.All`; deployment and tenant approval are separate and have not been performed. Graph consent/access and additional permissions required for hidden membership or constrained groups must be checked in the actual tenant. A configured permission request is not proof of an approved grant. See the [HTTP client guide](../hooks/http-clients.md) and [host checklist](../../SHAREPOINT-VALIDATION.md).

## Structural branches and lifetime

Place `Loading`, `Fallback` and `ErrorFallback` once each as **direct** children of their own family. They are structural content markers, not Fluent slots. The trim emits caller content without a DOM wrapper. Loading means a current read is pending; Fallback means all requested reads succeeded and access did not match; ErrorFallback means validation, transport, payload or resolution failed. Errors never select Fallback. Every omitted branch renders empty, including an omitted ErrorFallback. An absent `children` prop renders empty.

Inactive protected/branch components never render, run effects or mount. Caller DOM, class names, refs and formatting belong on the caller's children. Controls and markers expose no public DOM, class/style/ref or callback props. Changing the active role or the structural marker's React key remounts its content; the same role/key preserves caller state and child keys. Child arrays/keys remain meaningful. React 17's broad `ReactNode` type can admit functions through its historical `{}` alternative; callbacks are unsupported and never executed.

Duplicate, wrong-family and detectable fragment-wrapped markers fail closed. Keep markers direct rather than hiding them inside wrappers/fragments or custom components. Markers are supported only as direct children of their own trim family. A standalone marker renders its children unconditionally and performs no access check; standalone usage is unsupported. Custom components are ordinary protected children, not a way to declare a structural branch.

Provider, native client, PageContext/current-user identity and target changes hide previous asynchronous grants during render. Obsolete/unmounted requests cannot restore them. Identical inflight reads may be shared within the matching provider only; settled grants are not cached and there is no polling. Host changes still require a React render. Run/repeat in the sample deliberately changes the trim key to perform a fresh check. UI trimming is advisory; each server operation still needs authorization.

The lazy **Entra Group Trim** sample uses this actual control and native APIs, editable selectors, enum fields, all three branches, a protected counter and effect mount gauge. Local fixtures replace only unavailable transports and do not establish authenticated tenant behavior. See [development verification](../../DEVELOPMENT.md) and the [controls index](./INDEX.md).

`EntraGroupTrimBranchProps` describes each structural marker. `EntraGroupTrimProps` describes the control.

Graph paging currently accepts HTTPS next links on `graph.microsoft.com` only. National-cloud or other next-link hosts select ErrorFallback; those endpoints are not supported by this adapter.
