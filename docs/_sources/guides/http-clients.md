# Choose and invoke an HTTP client

Choose the HTTP client for the target resource and distinguish initialization, invocation state, request failure and authorization.

> Hooks for accessing SPFx HTTP clients with state management

## Overview

The four request-client hooks provide SPFx HTTP clients with invocation loading/error tracking. The token-provider hook exposes provider initialization state, and the permission-precheck hook exposes diagnostic state; their contracts are described separately below.

### When to Use Which Client

| Client | Use Case |
|--------|----------|
| `HttpClient` | External APIs, webhooks, public endpoints |
| `SPHttpClient` | SharePoint REST API (`/_api/`) |
| `AadHttpClient` | Custom Azure AD secured APIs |
| `AadTokenProvider` | Token acquisition and delegated API permission diagnostics |
| `useSPFxApiPermissionPrecheck` | Preflight Graph or custom API delegated scopes before enabling a feature |
| `MSGraphClient` | Microsoft Graph API |

### Initialization and requests

Client acquisition and request invocation are separate phases. Link the owning API for exact result members/defaults and catch action promises. Direct client calls bypass hook invocation state. The token-provider hook does not request a token or supply invocation tracking; the permission-precheck hook returns diagnostics rather than a request client.

## Concurrent invocations

The HTTP, SharePoint REST, AadHttpClient and Graph client hooks use the same invocation tracker. `isLoading` remains true while any invocation on the current client is pending. Starting an invocation clears displayed error; only the latest-started invocation can publish an error. Client replacement and unmount ignore obsolete UI updates while caller promises still resolve/reject. `clearError()` does not cancel requests. Initialization flags/errors are separate from invocation state.



## Configure API permission requests

For a custom AAD API, register its application, request its exposed scope in the consumer's `package-solution.json`, and have the tenant administrator approve the request. The token-acquisition resource and the request URL have separate roles; use `setResourceUrl` for later resource changes and wait for the replacement client.

```json
{
  "webApiPermissionRequests": [
    { "resource": "contoso-api", "scope": "user_impersonation" },
    { "resource": "Microsoft Graph", "scope": "User.Read" },
    { "resource": "Microsoft Graph", "scope": "Mail.Read" }
  ]
}
```

Declaring a request is not a grant. Choose least-required scopes for the actual feature and verify approval/server behavior in the tenant. Consumers wanting Microsoft's Graph TypeScript models can add `@microsoft/microsoft-graph-types` as a development dependency; the examples below use explicit local expected-response shapes without adding a package dependency. Neither typing option validates remote JSON.

For Graph JSON persistence and temporary download URLs, see [storage](storage.md#temporary-browser-download-urls).

## Access trim consumers

[SharePointPermissionTrim](../../api/controls/sharepoint-permission-trim.md) and [SharePointGroupTrim](../../api/controls/sharepoint-group-trim.md) reuse the provider's native SPHttpClient. [EntraGroupTrim](../../api/controls/entra-group-trim.md) reuses MSGraphClientV3. Their samples invoke actual native reads; configured User.Read and GroupMember.Read.All requests need separate tenant approval and do not prove hidden membership completeness. Client/provider identity changes invalidate prior asynchronous decisions. See the [host checklist](../development/sharepoint-validation/INDEX.md).

<!-- toolkit:include symbol="useSPFxHttpClient" section="example" title="Weather with readiness failure and retry" -->

<!-- toolkit:include symbol="useSPFxHttpClient" section="example" title="Post a webhook with invocation state" -->

<!-- toolkit:include symbol="useSPFxAadHttpClient" section="example" title="Custom API with initialization and request states" -->

<!-- toolkit:include symbol="useSPFxSPHttpClient" section="example" title="List items CRUD with visible actions" -->

<!-- toolkit:include symbol="useSPFxSPHttpClient" section="example" title="Cross-site request without stale base state" -->

<!-- toolkit:include symbol="useSPFxAadTokenProvider" section="example" title="Inspect delegated scopes from a manual token" -->

<!-- toolkit:include symbol="useSPFxMSGraphClient" section="example" title="User profile with initialization request and retry" -->

<!-- toolkit:include symbol="useSPFxMSGraphClient" section="example" title="Read recent high-importance mail" -->

<!-- toolkit:include symbol="useSPFxMSGraphClient" section="example" title="Read the coming week calendar view" -->

<!-- toolkit:include symbol="useSPFxAadHttpClient" section="example" title="Switch resource before acquiring another API client" -->
