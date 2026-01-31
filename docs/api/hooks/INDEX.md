# Hooks API Reference

> Complete reference for all SPFx React Toolkit hooks

## Overview

The toolkit provides **35 hooks** organized into functional categories. All hooks must be used within components wrapped by an [SPFx Provider](../core/providers.md).

## Categories

| Category | Hooks | Description |
|----------|-------|-------------|
| [Context](./context.md) | 3 | Core SPFx context and services |
| [Properties & Display](./properties.md) | 3 | Web part properties and display mode |
| [HTTP Clients](./http-clients.md) | 3 | SharePoint, Graph, and Azure AD APIs |
| [PnPjs](./pnpjs.md) | 2 | PnP/sp and PnP/graph instances |
| [UI & Theming](./theming.md) | 4 | Theme, Fluent UI 9, and container info |
| [User & Site](./user-site.md) | 5 | User, site, hub, and list information |
| [Environment](./environment.md) | 4 | Environment detection, Teams, locale |
| [Storage](./storage.md) | 3 | LocalStorage, SessionStorage, OneDrive |
| [Permissions](./permissions.md) | 2 | Permission checking |
| [Performance & Diagnostics](./performance.md) | 4 | Logging, timing, correlation |

---

## Quick Reference

### Context Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxContext()` | Full SPFx context object | [View](./context.md#usespfxcontext) |
| `useSPFxPageContext()` | Page context information | [View](./context.md#usespfxpagecontext) |
| `useSPFxInstanceInfo()` | Web part instance details | [View](./context.md#usespfxinstanceinfo) |

### Properties & Display Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxProperties<T>()` | Web part properties | [View](./properties.md#usespfxproperties) |
| `useSPFxDisplayMode()` | Edit/Read display mode | [View](./properties.md#usespfxdisplaymode) |
| `useSPFxPropertyPane()` | Property pane controls | [View](./properties.md#usespfxpropertypane) |

### HTTP Client Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxHttpClient()` | Generic HTTP requests | [View](./http-clients.md#usespfxhttpclient) |
| `useSPFxMSGraphClient()` | Microsoft Graph API | [View](./http-clients.md#usespfxmsgraphclient) |
| `useSPFxAadHttpClient(resourceId)` | Azure AD protected APIs | [View](./http-clients.md#usespfxaadhttpclient) |

### PnPjs Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxPnPSP()` | PnPjs SP instance | [View](./pnpjs.md#usespfxpnpsp) |
| `useSPFxPnPGraph()` | PnPjs Graph instance | [View](./pnpjs.md#usespfxpnpgraph) |

### UI & Theming Hooks

| Hook | Description | Docs |
|------|-------------|------|
| `useSPFxThemeInfo()` | SharePoint theme | [View](./theming.md#usespfxthemeinfo) |
| `useSPFxFluent9ThemeInfo()` | Fluent UI 9 theme | [View](./theming.md#usespfxfluent9themeinfo) |
| `useSPFxContainerSize()` | Container dimensions | [View](./theming.md#usespfxcontainersize) |
| `useSPFxContainerInfo()` | Section background info | [View](./theming.md#usespfxcontainerinfo) |

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
| `useSPFxLocalStorage<T>(key, default)` | Persistent storage | [View](./storage.md#usespfxlocalstorage) |
| `useSPFxSessionStorage<T>(key, default)` | Session storage | [View](./storage.md#usespfxsessionstorage) |
| `useSPFxOneDriveAppData<T>(file, default)` | Cloud storage | [View](./storage.md#usespfxonedriveappdata) |

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
| `useSPFxTenantProperty(key)` | Tenant properties | [View](./performance.md#usespfxtenantproperty) |

---

## Usage Pattern

All hooks follow the same basic pattern:

```tsx
import { 
  SPFxWebPartProvider, 
  useSPFxContext, 
  useSPFxProperties,
  // ... other hooks
} from '@apvee/spfx-react-toolkit';

// In your web part render method:
public render(): void {
  const element = (
    <SPFxWebPartProvider context={this.context}>
      <MyComponent />
    </SPFxWebPartProvider>
  );
  ReactDom.render(element, this.domElement);
}

// In your component:
const MyComponent: React.FC = () => {
  // Use any hooks here
  const ctx = useSPFxContext();
  const { title } = useSPFxProperties<IMyProps>();
  const { displayName } = useSPFxUserInfo();
  
  return <div>Hello {displayName}!</div>;
};
```

---

## See Also

- [Core Providers](../core/providers.md) - Provider components
- [Core Types](../core/types.md) - Type definitions
- [Introduction](../../INTRODUCTION.md) - Getting started

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
