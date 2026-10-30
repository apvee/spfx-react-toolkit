# Choose and invoke an HTTP client

> Hooks for accessing SPFx HTTP clients with state management

## Overview

These hooks provide access to SPFx HTTP clients for making API calls with automatic loading and error state management.

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

All hooks follow a consistent pattern:

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
