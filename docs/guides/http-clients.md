# Choose and invoke an HTTP client

> Hooks for accessing SPFx HTTP clients with state management

## Overview

The four request-client hooks provide SPFx HTTP clients with invocation loading/error tracking. The token-provider hook exposes provider initialization state, and the permission-precheck hook exposes diagnostic state; their contracts are described separately below.

| Hook | Returns | Init Type | Description |
|------|---------|-----------|-------------|
| [`useSPFxHttpClient`](../api/hooks/http-clients/http-client.md#usespfxhttpclient) | `SPFxHttpClientInfo` | Sync | Generic HTTP client for external APIs |
| [`useSPFxSPHttpClient`](../api/hooks/http-clients/sharepoint-http-client.md#usespfxsphttpclient) | `SPFxSPHttpClientInfo` | Sync | SharePoint REST API client |
| [`useSPFxAadHttpClient`](../api/hooks/http-clients/aad-http-client.md#usespfxaadhttpclient) | `SPFxAadHttpClientInfo` | Async | Azure AD secured API client |
| [`useSPFxAadTokenProvider`](../api/hooks/http-clients/aad-token-provider.md#usespfxaadtokenprovider) | `SPFxAadTokenProviderInfo` | Async | SPFx AAD token provider access |
| [`useSPFxApiPermissionPrecheck`](../api/hooks/http-clients/api-permission-precheck.md#usespfxapipermissionprecheck) | `SPFxApiPermissionPrecheckResult` | Async | Delegated Graph and custom API permission precheck |
| [`useSPFxMSGraphClient`](../api/hooks/http-clients/graph-client.md#usespfxmsgraphclient) | `SPFxMSGraphClientInfo` | Async | Microsoft Graph client |

### When to Use Which Client

| Client | Use Case |
|--------|----------|
| `HttpClient` | External APIs, webhooks, public endpoints |
| `SPHttpClient` | SharePoint REST API (`/_api/`) |
| `AadHttpClient` | Custom Azure AD secured APIs |
| `AadTokenProvider` | Token acquisition and delegated API permission diagnostics |
| `useSPFxApiPermissionPrecheck` | Preflight Graph or custom API delegated scopes before enabling a feature |
| `MSGraphClient` | Microsoft Graph API |

### State Management Pattern

`useSPFxHttpClient`, `useSPFxSPHttpClient`, `useSPFxAadHttpClient` and `useSPFxMSGraphClient` follow this request-client pattern:

| Property | Sync Hooks | Async Hooks | Description |
|----------|------------|-------------|-------------|
| `client` | `T \| undefined` | `T \| undefined` | Native SPFx client |
| `invoke` | ✅ | ✅ | Execute with state tracking |
| `isLoading` | ✅ | ✅ | Loading state during `invoke()` |
| `error` | ✅ | ✅ | Last error from `invoke()` |
| `clearError` | ✅ | ✅ | Clear error state |
| `isReady` | ✅ | ✅ | Client ready for use |
| `isInitializing` | ❌ | ✅ | Client being initialized |
| `initError` | ❌ | ✅ | Initialization error |

`useSPFxAadTokenProvider` exposes only `tokenProvider`, `isInitializing`, `initError` and `isReady`. It does not acquire a token itself or provide `client`, `invoke`, `isLoading`, `error` or `clearError`; callers own loading and failure handling for requests made through the returned provider.

`useSPFxApiPermissionPrecheck` returns permission-specific state instead of a client. Its result includes `configurationState`, `available`, `missing`, `warnings`, `unknown`, `check()`, `retry()`, and `retryWithoutCache()`.

---

## Concurrent invocations

The HTTP, SharePoint REST, AadHttpClient and Graph client hooks use the same invocation tracker. `isLoading` remains true while any invocation on the current client is pending. Starting an invocation clears displayed error; only the latest-started invocation can publish an error. Client replacement and unmount ignore obsolete UI updates while caller promises still resolve/reject. `clearError()` does not cancel requests. Initialization flags/errors are separate from invocation state.



## API references

For JSON file persistence through Graph, see the [OneDrive app-data hook](../api/hooks/storage/onedrive-app-data.md).

- [useSPFxHttpClient](../api/hooks/http-clients/http-client.md).
- [useSPFxSPHttpClient](../api/hooks/http-clients/sharepoint-http-client.md).
- [useSPFxAadHttpClient](../api/hooks/http-clients/aad-http-client.md).
- [useSPFxAadTokenProvider](../api/hooks/http-clients/aad-token-provider.md).
- [useSPFxApiPermissionPrecheck](../api/hooks/http-clients/api-permission-precheck.md).
- [useSPFxMSGraphClient](../api/hooks/http-clients/graph-client.md).

## Access trim consumers

[SharePointPermissionTrim](../api/controls/sharepoint-permission-trim.md) and [SharePointGroupTrim](../api/controls/sharepoint-group-trim.md) reuse the provider's native SPHttpClient. [EntraGroupTrim](../api/controls/entra-group-trim.md) reuses MSGraphClientV3. Their samples invoke actual native reads; configured User.Read and GroupMember.Read.All requests need separate tenant approval and do not prove hidden membership completeness. Client/provider identity changes invalidate prior asynchronous decisions. See the [host checklist](../development/sharepoint-validation/INDEX.md).
