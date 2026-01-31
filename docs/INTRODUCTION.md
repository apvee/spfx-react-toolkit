# SPFx React Toolkit

> A comprehensive React runtime and hooks library for SharePoint Framework (SPFx) with 35+ type-safe hooks

## Overview

**SPFx React Toolkit** is a production-ready library that simplifies SharePoint Framework development by providing a unified React context provider and a comprehensive collection of strongly-typed hooks. Built on [Jotai](https://jotai.org/) atomic state management, it delivers per-instance state isolation, automatic synchronization, and an ergonomic React Hooks API.

### Key Benefits

- **💪 Type-Safe**: Full TypeScript support with zero `any` usage
- **⚡ Optimized**: Jotai atomic state with per-instance scoping
- **🔄 Auto-Sync**: Bidirectional synchronization between React and SPFx
- **🎨 Universal**: Works with all SPFx component types
- **📦 Modular**: Tree-shakeable, minimal bundle impact

## Installation

```bash
npm install @apvee/spfx-react-toolkit
```

### Peer Dependencies

All peer dependencies are installed automatically with npm 7+:

| Dependency | Size | Purpose |
|------------|------|---------|
| **Jotai** | ~3KB | Core state management |
| **PnPjs** | 30-50KB | SharePoint API (tree-shakeable) |

## Quick Start

### 1. Wrap Your Component with a Provider

Choose the appropriate provider for your SPFx component type:

```tsx
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';
import * as ReactDom from 'react-dom';

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  public render(): void {
    const element = React.createElement(
      SPFxWebPartProvider,
      { instance: this },
      React.createElement(MyComponent)
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

### 2. Use Hooks in Your Components

```tsx
import { 
  useSPFxPageContext, 
  useSPFxProperties, 
  useSPFxDisplayMode 
} from '@apvee/spfx-react-toolkit';

const MyComponent: React.FC = () => {
  const pageContext = useSPFxPageContext();
  const { properties, setProperties } = useSPFxProperties<IMyWebPartProps>();
  const { isEdit } = useSPFxDisplayMode();

  return (
    <div>
      <h1>{properties?.title ?? 'Welcome'}</h1>
      <p>Site: {pageContext.web.title}</p>
      {isEdit && (
        <input 
          value={properties?.title ?? ''} 
          onChange={(e) => setProperties({ title: e.target.value })}
        />
      )}
    </div>
  );
};
```

## Providers

| Provider | SPFx Component Type |
|----------|---------------------|
| `SPFxWebPartProvider` | WebParts |
| `SPFxApplicationCustomizerProvider` | Application Customizers |
| `SPFxFieldCustomizerProvider` | Field Customizers |
| `SPFxListViewCommandSetProvider` | ListView Command Sets |

## Hooks by Category

### Context & Metadata
| Hook | Description |
|------|-------------|
| [`useSPFxContext`](./api/hooks/context.md#usespfxcontext) | Core SPFx context access |
| [`useSPFxPageContext`](./api/hooks/context.md#usespfxpagecontext) | SharePoint page context |
| [`useSPFxServiceScope`](./api/hooks/context.md#usespfxservicescope) | SPFx service scope for DI |
| [`useSPFxInstanceInfo`](./api/hooks/context.md#usespfxinstanceinfo) | Component instance metadata |

### Properties & Display
| Hook | Description |
|------|-------------|
| [`useSPFxProperties`](./api/hooks/properties.md#usespfxproperties) | Bidirectional property management |
| [`useSPFxDisplayMode`](./api/hooks/properties.md#usespfxdisplaymode) | Read/Edit mode detection |

### HTTP Clients
| Hook | Description |
|------|-------------|
| [`useSPFxHttpClient`](./api/hooks/http-clients.md#usespfxhttpclient) | Generic HTTP client |
| [`useSPFxSPHttpClient`](./api/hooks/http-clients.md#usespfxsphttpclient) | SharePoint REST API client |
| [`useSPFxAadHttpClient`](./api/hooks/http-clients.md#usespfxaadhttpclient) | Azure AD secured API client |
| [`useSPFxMSGraphClient`](./api/hooks/http-clients.md#usespfxmsgraphclient) | Microsoft Graph client |

### PnPjs Integration
| Hook | Description |
|------|-------------|
| [`useSPFxPnP`](./api/hooks/pnpjs.md#usespfxpnp) | PnPjs with state management |
| [`useSPFxPnPContext`](./api/hooks/pnpjs.md#usespfxpnpcontext) | PnPjs SPFI factory |
| [`useSPFxPnPList`](./api/hooks/pnpjs.md#usespfxpnplist) | List CRUD operations |
| [`useSPFxPnPSearch`](./api/hooks/pnpjs.md#usespfxpnpsearch) | SharePoint Search |

### UI & Theming
| Hook | Description |
|------|-------------|
| [`useSPFxThemeInfo`](./api/hooks/theming.md#usespfxthemeinfo) | SPFx theme (Fluent UI 8) |
| [`useSPFxFluent9ThemeInfo`](./api/hooks/theming.md#usespfxfluent9themeinfo) | Fluent UI 9 theme conversion |
| [`useSPFxContainerSize`](./api/hooks/theming.md#usespfxcontainersize) | Responsive breakpoints |
| [`useSPFxContainerInfo`](./api/hooks/theming.md#usespfxcontainerinfo) | Container dimensions |

### User & Site Information
| Hook | Description |
|------|-------------|
| [`useSPFxUserInfo`](./api/hooks/user-site.md#usespfxuserinfo) | Current user information |
| [`useSPFxUserPhoto`](./api/hooks/user-site.md#usespfxuserphoto) | User profile photo |
| [`useSPFxSiteInfo`](./api/hooks/user-site.md#usespfxsiteinfo) | Site and web information |
| [`useSPFxHubSiteInfo`](./api/hooks/user-site.md#usespfxhubsiteinfo) | Hub site information |
| [`useSPFxListInfo`](./api/hooks/user-site.md#usespfxlistinfo) | Current list context |

### Environment & Platform
| Hook | Description |
|------|-------------|
| [`useSPFxEnvironmentInfo`](./api/hooks/environment.md#usespfxenvironmentinfo) | Environment detection |
| [`useSPFxTeams`](./api/hooks/environment.md#usespfxteams) | Microsoft Teams context |
| [`useSPFxLocaleInfo`](./api/hooks/environment.md#usespfxlocaleinfo) | Locale and timezone |
| [`useSPFxPageType`](./api/hooks/environment.md#usespfxpagetype) | Page type detection |

### Storage
| Hook | Description |
|------|-------------|
| [`useSPFxLocalStorage`](./api/hooks/storage.md#usespfxlocalstorage) | Instance-scoped localStorage |
| [`useSPFxSessionStorage`](./api/hooks/storage.md#usespfxsessionstorage) | Instance-scoped sessionStorage |
| [`useSPFxOneDriveAppData`](./api/hooks/storage.md#usespfxonedriveappdata) | OneDrive app folder storage |

### Permissions
| Hook | Description |
|------|-------------|
| [`useSPFxPermissions`](./api/hooks/permissions.md#usespfxpermissions) | SharePoint permissions |
| [`useSPFxCrossSitePermissions`](./api/hooks/permissions.md#usespfxcrosssitepermissions) | Cross-site permissions |

### Performance & Diagnostics
| Hook | Description |
|------|-------------|
| [`useSPFxPerformance`](./api/hooks/performance.md#usespfxperformance) | Performance measurement |
| [`useSPFxLogger`](./api/hooks/performance.md#usespfxlogger) | Structured logging |
| [`useSPFxCorrelationInfo`](./api/hooks/performance.md#usespfxcorrelationinfo) | Request correlation |
| [`useSPFxTenantProperty`](./api/hooks/performance.md#usespfxtenantproperty) | Tenant properties |

## Requirements

| Requirement | Version |
|-------------|---------|
| Node.js | 22.x |
| SPFx | 1.18.0+ |
| React | 17.x |
| TypeScript | 5.3+ |

## Architecture

SPFx React Toolkit uses a layered architecture:

```
┌─────────────────────────────────────────┐
│           Your React Components          │
├─────────────────────────────────────────┤
│       35+ Type-Safe React Hooks          │
├─────────────────────────────────────────┤
│      Jotai Atoms (Instance-Scoped)       │
├─────────────────────────────────────────┤
│    SPFx Provider (Context + Sync)        │
├─────────────────────────────────────────┤
│         SPFx Runtime (WebPart, etc.)     │
└─────────────────────────────────────────┘
```

## License

MIT - See [LICENSE](../LICENSE) for details.

## Links

- [Full API Reference](./INDEX.md)
- [GitHub Repository](https://github.com/apvee/spfx-react-toolkit)
- [NPM Package](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
