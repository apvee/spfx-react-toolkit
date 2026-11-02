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

## Explore explicit client scenarios

Mount the source-owned scenarios below the [provider matching the native host](../../api/core/providers.md). They keep business reads behind explicit commands and can render independently or forward their regions to the sample explorer. Client/provider factory initialization is separate from a business request or an explicit token acquisition.

### Inspect readiness before choosing an operation

Inspect all five client/provider states without adding a token request. Graph client and token-provider acquisition can begin when their hooks mount; initialization errors and readiness remain separate from request loading and errors.

<!-- toolkit:include symbol="useSPFxAadTokenProvider" section="example" title="Inspect client readiness and token-provider initialization" -->

### Load external JSON

Use the external todo scenario to observe a tracked invocation, caught action failures and returned scalar fields. The HTTP, SharePoint and AAD JSON callbacks in these scenarios parse the response without inspecting `response.ok`; the hook does not turn a non-success HTTP response into a rejection automatically. Add the required status policy to your own callback.

<!-- toolkit:include symbol="useSPFxHttpClient" section="example" title="Load an external todo with tracked invocation state" -->

### Read the current SharePoint web

The current-web scenario synchronizes its request base through `setBaseUrl` when native page metadata changes, without issuing a read. The hook's initial base URL is captured state; it does not follow later page changes by itself.

<!-- toolkit:include symbol="useSPFxSPHttpClient" section="example" title="Read the current SharePoint web title and URL" -->

### Load the Graph profile after initialization

Wait for Graph acquisition to finish, then explicitly load `/me`. Keep initialization failures separate from invocation failures and verify the required delegated `User.Read` approval for the actual tenant.

<!-- toolkit:include symbol="useSPFxMSGraphClient" section="example" title="Read the current Graph profile after initialization" -->

### Select a custom API resource and call a path

Edit the resource and path drafts, initialize the trimmed resource explicitly, then call only when the acquired client matches that resource. Editing either draft invalidates the old local result. An `api://` identifier can identify a token resource without being an HTTP endpoint; the unchanged failed resource has no dedicated initialization-retry API.

<!-- toolkit:include symbol="useSPFxAadHttpClient" section="example" title="Initialize a custom API client and call a path" -->

## Concurrent invocations

The HTTP, SharePoint REST, AadHttpClient and Graph client hooks use the same invocation tracker. `isLoading` remains true while any invocation on the current client is pending. Starting an invocation clears displayed error; only the latest-started invocation can publish an error. Client replacement and unmount ignore obsolete UI updates while caller promises still resolve/reject. `clearError()` does not cancel requests. Initialization flags/errors are separate from invocation state.


## Reset local results and check delegated scopes

The executable request scenarios own result publication and invalidation in addition to the hook's invocation state. Their optional `resetRevision` starts at zero; its first mounted value is a baseline, and a later change restores the scenario's local drafts/results while retaining the mounted component and controls. Reset invalidates obsolete local completions and retained commands. It does not cancel native requests or initialization, rebuild the provider, clear token caches or grant permissions. The examples continue to show activity from older native requests even when a client selection changes; `isLoading` itself describes only invocations on the current client. A readiness snapshot keeps the actual native initialization state through Reset.

The manual permission scenario uses the mounted [useSPFxApiPermissionPrecheck hook](../../api/hooks/http-clients/api-permission-precheck.md) with `autoCheck: false` and `mode: 'passive'`. It retains only completed results returned by `check()` or `retryWithoutCache()` and derives their diagnostics through `summarizeSPFxApiPermissionResults`. A configuration/provider change can reset hook checking state before the original public promise settles, so the scenario keeps a separate original-promise latch and blocks another check until that promise settles. Reset clears its local completed snapshot and resource draft while preserving that pending truth.

An empty completed-results view summarizes to `configurationState: 'idle'` with `isConfigured: true`; this is not a completed grant check. Passive popup/redirect guards detach when the overall public check settles. A timeout does not cancel native token acquisition, which may settle later. Retry without cache forwards an uncached token preference while keeping passive mode; it neither clears a cache nor changes tenant grants. Token diagnostics do not establish server authorization or prove a later API request will succeed.

<!-- toolkit:include symbol="useSPFxApiPermissionPrecheck" section="example" title="Manually check Graph and custom API delegated scopes" -->

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
