# Hooks API Reference

> Complete reference for SPFx React Toolkit hooks

## Overview

Hooks are the React-facing API. They must be used within components wrapped by a host-specific SPFx provider such as `SPFxWebPartProvider`, `SPFxApplicationCustomizerProvider`, `SPFxFieldCustomizerProvider`, or `SPFxListViewCommandSetProvider`.

## Categories

| Category | Hooks | Description |
|----------|-------|-------------|
| [Context](./context.md) | 4 | Core SPFx context and service scope |
| [Properties & Display](./properties.md) | 3 | Web part properties and display mode |
| [HTTP Clients](./http-clients.md) | 6 | SharePoint, Graph, Azure AD APIs, token provider access, and API permission prechecks |
| [PnPjs](./pnpjs.md) | 4 | PnPjs context, invoke/batch, lists, and search |
| [UI & Theming](./theming.md) | 4 | Theme, Fluent UI 9, and container info |
| [User & Site](./user-site.md) | 5 | User, site, hub, and list information |
| [Environment](./environment.md) | 4 | Environment detection, Teams, locale, and page type |
| [Storage](./storage.md) | 5 | Browser, OneDrive, and tenant storage |
| [Permissions](./permissions.md) | 2 | Permission checking |
| [Performance & Diagnostics](./performance.md) | 3 | Logging, timing, correlation |

---

## Quick Reference

### Context Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxContext()` | Full SPFx context object | [View](./context.md#usespfxcontext) |
| `useSPFxPageContext()` | Page context information | [View](./context.md#usespfxpagecontext) |
| `useSPFxServiceScope()` | SPFx service scope access | [View](./context.md#usespfxservicescope) |
| `useSPFxInstanceInfo()` | Component instance details | [View](./context.md#usespfxinstanceinfo) |

### Properties & Display Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxProperties<T>()` | Component properties with bidirectional sync | [View](./properties.md#usespfxproperties) |
| `useSPFxDisplayMode()` | Edit/Read display mode | [View](./properties.md#usespfxdisplaymode) |
| `useSPFxIsEdit()` | Boolean shortcut for edit mode | [View](./properties.md#usespfxisedit) |

### HTTP Client Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxHttpClient()` | Generic HTTP requests | [View](./http-clients.md#usespfxhttpclient) |
| `useSPFxSPHttpClient()` | SharePoint REST requests | [View](./http-clients.md#usespfxsphttpclient) |
| `useSPFxMSGraphClient()` | Microsoft Graph API | [View](./http-clients.md#usespfxmsgraphclient) |
| `useSPFxAadHttpClient(resourceId)` | Azure AD protected APIs | [View](./http-clients.md#usespfxaadhttpclient) |
| `useSPFxAadTokenProvider()` | SPFx AAD token provider access | [View](./http-clients.md#usespfxaadtokenprovider) |
| `useSPFxApiPermissionPrecheck(config, options?)` | Delegated Graph and custom API permission precheck | [View](./http-clients.md#usespfxapipermissionprecheck) |

#### API Permission Precheck Example

```tsx
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';

function PermissionNotice() {
  const precheck = useSPFxApiPermissionPrecheck({
    graph: ['Sites.Read.All'],
    customApis: [
      {
        name: 'Orders API',
        resource: 'api://contoso-orders-api',
        packageResource: 'Orders API',
        scopes: ['Orders.Read']
      }
    ]
  });

  if (precheck.configurationState === 'ready') {
    return null;
  }

  return (
    <ul>
      {precheck.missing.map(item => (
        <li key={item.id}>{item.adminMessage}</li>
      ))}
    </ul>
  );
}
```

#### Manual Check Example

```tsx
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';

function ManualPermissionStatus() {
  const precheck = useSPFxApiPermissionPrecheck(
    { graph: ['Sites.Read.All'] },
    { autoCheck: false, mode: 'passive' }
  );

  return (
    <button onClick={() => precheck.retryWithoutCache()} disabled={precheck.isChecking}>
      Recheck permissions
    </button>
  );
}
```

`mode: 'passive'` prevents the precheck from launching an authentication popup or redirect while it probes token availability. Use `retryWithoutCache()` after an administrator changes API access or when you need SPFx to bypass a cached token on the next check.

### PnPjs Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxPnP()` | PnPjs invoke and batch helpers | [View](./pnpjs.md#usespfxpnp) |
| `useSPFxPnPContext()` | PnPjs `SPFI` factory | [View](./pnpjs.md#usespfxpnpcontext) |
| `useSPFxPnPList<T>()` | SharePoint list CRUD and batch operations | [View](./pnpjs.md#usespfxpnplist) |
| `useSPFxPnPSearch<T>()` | SharePoint Search with pagination | [View](./pnpjs.md#usespfxpnpsearch) |

### UI & Theming Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxThemeInfo()` | SharePoint theme | [View](./theming.md#usespfxthemeinfo) |
| `useSPFxFluent9ThemeInfo()` | Fluent UI 9 theme | [View](./theming.md#usespfxfluent9themeinfo) |
| `useSPFxContainerSize()` | Responsive container size category | [View](./theming.md#usespfxcontainersize) |
| `useSPFxContainerInfo()` | Container element and dimensions | [View](./theming.md#usespfxcontainerinfo) |

### User & Site Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxUserInfo()` | Current user | [View](./user-site.md#usespfxuserinfo) |
| `useSPFxUserPhoto(options?)` | User profile photo | [View](./user-site.md#usespfxuserphoto) |
| `useSPFxSiteInfo()` | Site collection info | [View](./user-site.md#usespfxsiteinfo) |
| `useSPFxHubSiteInfo()` | Hub site info | [View](./user-site.md#usespfxhubsiteinfo) |
| `useSPFxListInfo()` | Current list context | [View](./user-site.md#usespfxlistinfo) |

### Environment Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxEnvironmentInfo()` | Runtime environment | [View](./environment.md#usespfxenvironmentinfo) |
| `useSPFxTeams()` | Teams context | [View](./environment.md#usespfxteams) |
| `useSPFxLocaleInfo()` | Locale settings | [View](./environment.md#usespfxlocaleinfo) |
| `useSPFxPageType()` | Page type info | [View](./environment.md#usespfxpagetype) |

### Storage Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxLocalStorage<T>(key, defaultValue)` | Instance-scoped persistent browser storage | [View](./storage.md#usespfxlocalstorage) |
| `useSPFxSessionStorage<T>(key, defaultValue)` | Instance-scoped session browser storage | [View](./storage.md#usespfxsessionstorage) |
| `useSPFxOneDriveAppData<T>(fileName, defaultValue)` | OneDrive app folder storage | [View](./storage.md#usespfxonedriveappdata) |
| `useSPFxTenantProperty<T>(key)` | Tenant properties (read-only) | [View](./storage.md#usespfxtenantproperty) |
| `useSPFxTenantKeyValueStore()` | Tenant key-value store | [View](./storage.md#usespfxtenantkeyvaluestore) |

### Permissions Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxPermissions()` | Current site permissions | [View](./permissions.md#usespfxpermissions) |
| `useSPFxCrossSitePermissions(url)` | Cross-site permissions | [View](./permissions.md#usespfxcrosssitepermissions) |

### Performance & Diagnostics Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxPerformance()` | Performance timing | [View](./performance.md#usespfxperformance) |
| `useSPFxLogger()` | Structured logging | [View](./performance.md#usespfxlogger) |
| `useSPFxCorrelationInfo()` | Request correlation | [View](./performance.md#usespfxcorrelationinfo) |

---

## Usage Pattern

All hooks follow the same basic pattern:

```tsx
import {
  SPFxWebPartProvider,
  useSPFxContext,
  useSPFxProperties,
} from '@apvee/spfx-react-toolkit';

// In your web part render method:
public render(): void {
  const element = (
    <SPFxWebPartProvider instance={this}>
      <MyComponent />
    </SPFxWebPartProvider>
  );
  ReactDom.render(element, this.domElement);
}

// In your component:
const MyComponent: React.FC = () => {
  const ctx = useSPFxContext();
  const { properties } = useSPFxProperties<IMyProps>();
  return <div>{ctx.instanceId}: {properties?.title}</div>;
};
```

## See Also

- [Core Providers](../core/providers.md) - Provider components
- [Core Types](../core/types.md) - Type definitions
- [Helpers API](../helpers/INDEX.md) - Pure utility functions
- [Services API](../services/INDEX.md) - Non-React service factories
- [Introduction](../../INTRODUCTION.md) - Getting started
