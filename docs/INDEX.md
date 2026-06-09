# API Reference

> Complete API documentation for @apvee/spfx-react-toolkit

## Table of Contents

- [Core](#core)
  - [Providers](#providers)
  - [Types](#types)
- [Hooks](#hooks)
  - [Context & Metadata](#context--metadata)
  - [Properties & Display](#properties--display)
  - [HTTP Clients](#http-clients)
  - [PnPjs Integration](#pnpjs-integration)
  - [UI & Theming](#ui--theming)
  - [User & Site Information](#user--site-information)
  - [Environment & Platform](#environment--platform)
  - [Storage](#storage)
  - [Permissions](#permissions)
  - [Performance & Diagnostics](#performance--diagnostics)
- [Services](#services)
- [Helpers](#helpers)

---

## Core

### Providers

React context providers that wrap your SPFx components and enable hook usage.

| Provider | Component Type | Documentation |
|----------|----------------|---------------|
| `SPFxWebPartProvider` | WebParts | [providers.md](./api/core/providers.md#spfxwebpartprovider) |
| `SPFxApplicationCustomizerProvider` | Application Customizers | [providers.md](./api/core/providers.md#spfxapplicationcustomizerprovider) |
| `SPFxFieldCustomizerProvider` | Field Customizers | [providers.md](./api/core/providers.md#spfxfieldcustomizerprovider) |
| `SPFxListViewCommandSetProvider` | ListView Command Sets | [providers.md](./api/core/providers.md#spfxlistviewcommandsetprovider) |

### Types

Core TypeScript type definitions.

| Type | Description | Documentation |
|------|-------------|---------------|
| `HostKind` | SPFx component type discriminator | [types.md](./api/core/types.md#hostkind) |
| `SPFxComponent` | Union of all SPFx component types | [types.md](./api/core/types.md#spfxcomponent) |
| `SPFxContextType` | Union of all SPFx context types | [types.md](./api/core/types.md#spfxcontexttype) |
| `ContainerSize` | Container dimensions interface | [types.md](./api/core/types.md#containersize) |
| `SPFxProviderProps` | Provider component props | [types.md](./api/core/types.md#spfxproviderprops) |
| `SPFxContextValue` | Context value interface | [types.md](./api/core/types.md#spfxcontextvalue) |

---

## Hooks

### Context & Metadata

Access SPFx context and instance metadata.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxContext`](./api/hooks/context.md#usespfxcontext) | `SPFxContextValue` | Core SPFx context (instanceId, spfxContext, kind) |
| [`useSPFxPageContext`](./api/hooks/context.md#usespfxpagecontext) | `PageContext` | SharePoint page context |
| [`useSPFxServiceScope`](./api/hooks/context.md#usespfxservicescope) | `SPFxServiceScopeInfo` | Service scope for dependency injection |
| [`useSPFxInstanceInfo`](./api/hooks/context.md#usespfxinstanceinfo) | `SPFxInstanceInfo` | Component instance metadata |

### Properties & Display

Manage component properties and display state.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxProperties`](./api/hooks/properties.md#usespfxproperties) | `SPFxPropertiesInfo<T>` | Bidirectional property management |
| [`useSPFxDisplayMode`](./api/hooks/properties.md#usespfxdisplaymode) | `SPFxDisplayModeInfo` | Read/Edit mode detection |
| [`useSPFxIsEdit`](./api/hooks/properties.md#usespfxisedit) | `boolean` | Shortcut for edit mode check |

### HTTP Clients

Access SPFx HTTP clients for API calls.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxHttpClient`](./api/hooks/http-clients.md#usespfxhttpclient) | `SPFxHttpClientInfo` | Generic HTTP client |
| [`useSPFxSPHttpClient`](./api/hooks/http-clients.md#usespfxsphttpclient) | `SPFxSPHttpClientInfo` | SharePoint REST API client |
| [`useSPFxAadHttpClient`](./api/hooks/http-clients.md#usespfxaadhttpclient) | `SPFxAadHttpClientInfo` | Azure AD secured API client |
| [`useSPFxAadTokenProvider`](./api/hooks/http-clients.md#usespfxaadtokenprovider) | `SPFxAadTokenProviderInfo` | SPFx AAD token provider access |
| [`useSPFxApiPermissionPrecheck`](./api/hooks/http-clients.md#usespfxapipermissionprecheck) | `SPFxApiPermissionPrecheckResult` | Delegated Graph and custom API permission precheck |
| [`useSPFxMSGraphClient`](./api/hooks/http-clients.md#usespfxmsgraphclient) | `SPFxMSGraphClientInfo` | Microsoft Graph client |

### PnPjs Integration

PnPjs v4 integration with state management.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxPnP`](./api/hooks/pnpjs.md#usespfxpnp) | `SPFxPnPInfo` | PnPjs with invoke/batch helpers |
| [`useSPFxPnPContext`](./api/hooks/pnpjs.md#usespfxpnpcontext) | `PnPContextInfo` | PnPjs SPFI factory |
| [`useSPFxPnPList`](./api/hooks/pnpjs.md#usespfxpnplist) | `SPFxPnPListInfo<T>` | List CRUD operations |
| [`useSPFxPnPSearch`](./api/hooks/pnpjs.md#usespfxpnpsearch) | `SPFxPnPSearchInfo<T>` | SharePoint Search with pagination |

### UI & Theming

Access theme and layout information.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxThemeInfo`](./api/hooks/theming.md#usespfxthemeinfo) | `IReadonlyTheme \| undefined` | SPFx theme (Fluent UI 8) |
| [`useSPFxFluent9ThemeInfo`](./api/hooks/theming.md#usespfxfluent9themeinfo) | `SPFxFluent9ThemeInfo` | Fluent UI 9 theme conversion |
| [`useSPFxContainerSize`](./api/hooks/theming.md#usespfxcontainersize) | `SPFxContainerSizeInfo` | Responsive breakpoints |
| [`useSPFxContainerInfo`](./api/hooks/theming.md#usespfxcontainerinfo) | `SPFxContainerInfo` | Container dimensions |

### User & Site Information

Access user and site metadata.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxUserInfo`](./api/hooks/user-site.md#usespfxuserinfo) | `SPFxUserInfo` | Current user information |
| [`useSPFxUserPhoto`](./api/hooks/user-site.md#usespfxuserphoto) | `SPFxUserPhotoResult` | User profile photo URL |
| [`useSPFxSiteInfo`](./api/hooks/user-site.md#usespfxsiteinfo) | `SPFxSiteInfo` | Site and web information |
| [`useSPFxHubSiteInfo`](./api/hooks/user-site.md#usespfxhubsiteinfo) | `SPFxHubSiteInfo` | Hub site information |
| [`useSPFxListInfo`](./api/hooks/user-site.md#usespfxlistinfo) | `SPFxListInfo \| undefined` | Current list context |

### Environment & Platform

Detect environment and platform capabilities.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxEnvironmentInfo`](./api/hooks/environment.md#usespfxenvironmentinfo) | `SPFxEnvironmentInfo` | Environment type detection |
| [`useSPFxTeams`](./api/hooks/environment.md#usespfxteams) | `SPFxTeamsInfo` | Microsoft Teams context |
| [`useSPFxLocaleInfo`](./api/hooks/environment.md#usespfxlocaleinfo) | `SPFxLocaleInfo` | Locale and timezone |
| [`useSPFxPageType`](./api/hooks/environment.md#usespfxpagetype) | `SPFxPageTypeInfo` | Page type detection |

### Storage

Persistent storage with instance scoping.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxLocalStorage`](./api/hooks/storage.md#usespfxlocalstorage) | `SPFxStorageHook<T>` | Instance-scoped localStorage |
| [`useSPFxSessionStorage`](./api/hooks/storage.md#usespfxsessionstorage) | `SPFxStorageHook<T>` | Instance-scoped sessionStorage |
| [`useSPFxOneDriveAppData`](./api/hooks/storage.md#usespfxonedriveappdata) | `SPFxOneDriveAppDataResult<T>` | OneDrive app folder storage |
| [`useSPFxTenantProperty`](./api/hooks/storage.md#usespfxtenantproperty) | `SPFxTenantPropertyResult<T>` | Tenant properties (read-only) |
| [`useSPFxTenantKeyValueStore`](./api/hooks/storage.md#usespfxtenantkeyvaluestore) | `SPFxTenantKeyValueStoreResult` | Tenant-level key-value store |

### Permissions

SharePoint permissions checking.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxPermissions`](./api/hooks/permissions.md#usespfxpermissions) | `SPFxPermissionsInfo` | Site/web/list permissions |
| [`useSPFxCrossSitePermissions`](./api/hooks/permissions.md#usespfxcrosssitepermissions) | `SPFxCrossSitePermissionsInfo` | Cross-site permissions |

### Performance & Diagnostics

Performance measurement and logging.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxPerformance`](./api/hooks/performance.md#usespfxperformance) | `SPFxPerformanceInfo` | Performance measurement |
| [`useSPFxLogger`](./api/hooks/performance.md#usespfxlogger) | `SPFxLoggerInfo` | Structured logging |
| [`useSPFxCorrelationInfo`](./api/hooks/performance.md#usespfxcorrelationinfo) | `SPFxCorrelationInfo` | Request correlation IDs |

---

## Services

Public non-React service factories for composing SPFx, PnPjs, SharePoint REST, and Microsoft Graph operations outside hooks.

| Module | Description | Documentation |
|--------|-------------|---------------|
| PnPjs services | Context, invoke/batch, list CRUD, and search | [services/INDEX.md](./api/services/INDEX.md) |
| Tenant services | App catalog discovery, tenant properties, tenant key-value store | [services/INDEX.md](./api/services/INDEX.md) |
| Graph services | OneDrive app data and user photos | [services/INDEX.md](./api/services/INDEX.md) |
| API permission precheck service | Token-provider-based delegated scope checks with remediation entries for `webApiPermissionRequests` | [services/INDEX.md](./api/services/INDEX.md) |

---

## Helpers

Public pure helper functions for SPFx context mapping, permission checks, API permission precheck normalization and summarization, storage keys, tenant value serialization, Graph paths, and theme conversion.

| Module | Description | Documentation |
|--------|-------------|---------------|
| Page context helpers | User, site, list, locale, environment, page type, and correlation mapping | [helpers/INDEX.md](./api/helpers/INDEX.md) |
| Utility helpers | Permissions, container size, storage keys, tenant values, Graph paths, theme conversion | [helpers/INDEX.md](./api/helpers/INDEX.md) |
| API permission precheck helpers | Requirement normalization, JWT scope evaluation, error classification, and summaries that map remediation to `webApiPermissionRequests` | [helpers/INDEX.md](./api/helpers/INDEX.md) |

---

## Quick Links

- [Introduction & Quick Start](./INTRODUCTION.md)
- [Core Module](./api/core/INDEX.md)
- [Hooks Module](./api/hooks/INDEX.md)
- [Services Module](./api/services/INDEX.md)
- [Helpers Module](./api/helpers/INDEX.md)

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
