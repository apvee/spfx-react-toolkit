# SPFx React Toolkit

> React controls, four host providers, 46 type-safe hooks, helpers and services for SharePoint Framework (SPFx)

## Overview

**SPFx React Toolkit** is a library that simplifies SharePoint Framework development by providing a unified React context provider and a comprehensive collection of strongly-typed hooks. Built on a provider-scoped runtime store, it delivers per-instance state isolation, automatic synchronization, and an ergonomic React Hooks API.

### Key Benefits

- **💪 Type-Safe**: Public TypeScript declarations for hooks, helpers and services
- **⚡ Optimized**: Provider-scoped runtime state with per-instance isolation
- **🔄 Auto-Sync**: Bidirectional synchronization between React and SPFx
- **🎨 Host providers**: Four exported providers for their matching SPFx component types
- **📦 Modular**: ESNext library modules compiled independently from the sample app

## Installation

```bash
npm install @apvee/spfx-react-toolkit
```

### Peer Dependencies

The consuming SPFx project supplies compatible React, ReactDOM and SPFx runtime packages. PnPjs hooks/services use `@pnp/core`, `@pnp/queryable` and `@pnp/sp` v4 peers. Check the host's dependency versions rather than relying on npm to choose a newer SPFx toolchain.

Fluent and Griffel integration requires shared peers `@fluentui/react-components` (`^9.74.9`), `@fluentui/react-charts` (`^9.3.27`), `@griffel/core` (`^1.21.4`), `@griffel/react` (`^1.7.8`), `@fluentui/react-shared-contexts` (`^9.26.4`), `@fluentui/react-migration-v8-v9` (`^9.10.20`), `@fluentui/react-theme` (`^9.2.2`) and `@fluentui/react-utilities` (`^9.26.7`). Token-based `useSx` selections need theme variables in scope, usually from the host’s existing FluentProvider. The optional provider package used by the sample is `@fluentui/react-provider@9.22.8`; no new provider is required by useSx.

The package also declares `react-hook-form` (`>=7.87.0 <8.0.0`), `@fluentui/react-datepicker-compat` (`>=0.6.16 <1.0.0`), `@fluentui/react-timepicker-compat` (`>=0.4.21 <1.0.0`) and `@fluentui/react-icons` (`>=2.0.312 <3.0.0`) as mandatory peers. The native Property Pane factories use the host's `@microsoft/sp-property-pane` (`>=1.18.0 <2.0.0`). Modern npm can install missing peers automatically; keep the host's compatible versions and shared runtime identities. See [Form dependencies](./api/controls/form.md#dependencies) for the tested baseline.

### Register localized UI resources

In the consuming SPFx application's `package.json`, add `"postinstall": "spfx-react-toolkit configure-localization"` (or append the command to its existing postinstall), then run once from that application:

```bash
npx --no-install spfx-react-toolkit configure-localization
```

This registers the eight native resource IDs in `config/config.json`. Typed `strings` overrides merge with those resources and still require their SPFx loader mappings. See [Localization](./LOCALIZATION.md) for monorepo setup, checks, culture selection and the removed `EditableText.locale` prop.

The repository is an npm workspace monorepo: [library source](../packages/spfx-react-toolkit/src/index.ts) builds with TypeScript to ESNext `lib`; the private [SPFx sample](../apps/spfx-react-toolkit-test/package.json) bundles with SPFx 1.21.1 Gulp. See [Development](./DEVELOPMENT.md) and [SharePoint validation](./SHAREPOINT-VALIDATION.md). Library changes require a library rebuild and a serve restart.

## Generic React styles

`const sx = useSx()` returns a class composer: `sx(flex.row, gap.medium, presets.canvas, padding.medium)`. Import these names from the package root and call the hook inside a component. See [Styles and useSx](./api/helpers/styles.md) and [React utility hooks](./api/hooks/react.md) for the full contracts. The lazy Styles sample panel exercises layout, catalog selections, queries, states and scrolling.

## Controls and Property Pane fields

The [controls reference](./api/controls/INDEX.md) covers text/editing, loading and state messages, settings/disclosure groups, Form, DataGrid, Stage, navigation, metadata lists, cards, access trims and SharePoint pickers. Use `WebPartShell` with an actual web part instance for integrated providers and an editable title. Generic controls document their own provider and resource requirements.

Native [Property Pane pickers](./api/property-pane/pickers.md) and [file/folder picker](./api/property-pane/file-folder-picker.md) use dedicated `/property-pane/...` imports. Their factories are not root value exports.

## Quick Start

### 1. Wrap Your Component with a Provider

Choose the appropriate provider for your SPFx component type:

```tsx
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

interface IMyWebPartProps {
  title: string;
}

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
import * as React from 'react';
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
| [`useSPFxIsEdit`](./api/hooks/properties.md#usespfxisedit) | Boolean shortcut for edit mode |

### HTTP Clients
| Hook | Description |
|------|-------------|
| [`useSPFxHttpClient`](./api/hooks/http-clients.md#usespfxhttpclient) | Generic HTTP client |
| [`useSPFxSPHttpClient`](./api/hooks/http-clients.md#usespfxsphttpclient) | SharePoint REST API client |
| [`useSPFxAadHttpClient`](./api/hooks/http-clients.md#usespfxaadhttpclient) | Azure AD secured API client |
| [`useSPFxAadTokenProvider`](./api/hooks/http-clients.md#usespfxaadtokenprovider) | SPFx AAD token provider access |
| [`useSPFxApiPermissionPrecheck`](./api/hooks/http-clients.md#usespfxapipermissionprecheck) | Delegated Graph and custom API permission precheck |
| [`useSPFxMSGraphClient`](./api/hooks/http-clients.md#usespfxmsgraphclient) | Microsoft Graph client |

### PnPjs Integration
| Hook | Description |
|------|-------------|
| [`useSPFxPnP`](./api/hooks/pnpjs.md#usespfxpnp) | PnPjs with state management |
| [`useSPFxPnPContext`](./api/hooks/pnpjs.md#usespfxpnpcontext) | PnPjs SPFI factory |
| [`useSPFxPnPList`](./api/hooks/pnpjs.md#usespfxpnplist) | List CRUD operations |
| [`useSPFxPnPListById`](./api/hooks/pnpjs.md#usespfxpnplistbyid) | List access by GUID |
| [`useSPFxPnPListByUrl`](./api/hooks/pnpjs.md#usespfxpnplistbyurl) | List access by decoded server-relative root URL |
| [`useSPFxPnPListByPath`](./api/hooks/pnpjs.md#usespfxpnplistbypath) | List access by decoded web-relative root path |
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
| [`useSPFxTenantProperty`](./api/hooks/storage.md#usespfxtenantproperty) | Tenant properties (read-only) |
| [`useSPFxTenantKeyValueStore`](./api/hooks/storage.md#usespfxtenantkeyvaluestore) | Tenant-level key-value store |
| [`useSPFxSiteKeyValueStore`](./api/hooks/storage.md#usespfxsitekeyvaluestore) | Collection root-web key-value store shared by subsites |
| [`useSPFxConfigurationStore`](./api/hooks/storage.md#usespfxconfigurationstore) | Tenant defaults, site overrides and explicitly scoped writes |

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

## Helpers And Services

Hooks are the React API. Helpers and services are public non-React APIs for composition outside components or for building custom abstractions.

### Helpers

Helpers do not send network requests or read provider state. Most map only their supplied inputs; API permission token evaluation also checks the current clock, and theme helpers can return shared immutable objects. The styles facade additionally exports the React `useSx` hook, whose composer uses the existing renderer to insert CSS.

| Helper group | Examples | Documentation |
|--------------|----------|---------------|
| Page context mapping | `getSPFxUserInfo`, `getSPFxSiteInfo`, `getSPFxEnvironmentInfo` | [Helpers API](./api/helpers/INDEX.md) |
| Utility helpers | `createScopedSPFxStorageKey`, `getSPFxContainerSize`, `hasSPFxPermission` | [Helpers API](./api/helpers/INDEX.md) |
| API permission precheck helpers | `normalizeSPFxApiPermissionRequirements`, `summarizeSPFxApiPermissionResults` | [Helpers API](./api/helpers/INDEX.md) |
| Graph and theme helpers | `buildOneDriveAppDataPath`, `buildUserPhotoEndpoint`, `createFluent9ThemeFromSPFxTheme` | [Helpers API](./api/helpers/INDEX.md) |

### Services

Services perform reusable I/O operations with dependencies supplied by the caller.

| Service group | Examples | Documentation |
|---------------|----------|---------------|
| PnPjs services | `createSPFxPnPContextService`, `createSPFxPnPListService`, `createSPFxPnPSearchService` | [Services API](./api/services/INDEX.md) |
| Tenant services | `createSPFxAppCatalogService`, `createSPFxTenantPropertyService`, `createSPFxTenantKeyValueStoreService` | [Services API](./api/services/INDEX.md) |
| Site collection service | `createSPFxSiteKeyValueStoreService` | [Services API](./api/services/INDEX.md#createspfxsitekeyvaluestoreservice) |
| Configuration service | `createSPFxConfigurationStoreService` | [Services API](./api/services/INDEX.md#configuration-store-service) |
| Graph services | `createSPFxOneDriveAppDataService`, `createSPFxUserPhotoService` | [Services API](./api/services/INDEX.md) |
| API permission precheck services | `createSPFxApiPermissionPrecheckService` | [Services API](./api/services/INDEX.md) |

API permission precheck helpers, services, and hooks evaluate delegated token scopes for Graph and custom APIs. Remediation messages map each requirement back to the `webApiPermissionRequests` resource/scope pair that an administrator reviews in SharePoint admin center.

## Requirements

| Requirement | Repository verification baseline |
|-------------|----------------------------------|
| Node.js | `>=22.14.0 <23.0.0` |
| SPFx build/runtime packages | `1.21.1` |
| React / ReactDOM | `17.0.1` (package peers: `17.x`) |
| TypeScript | `5.3.3` |
| PnPjs | `4.17.0` (package peers: `^4.0.0`) |

The package declares SPFx peers `>=1.18.0 <2.0.0`; that range is a compatibility declaration, not evidence of testing every version. Repository checks use SPFx 1.21.1. Authenticated tenant validation is a separate step.

## License

MIT - See [LICENSE](../LICENSE) for details.

## Links

- [Full API Reference](./INDEX.md)
- [GitHub Repository](https://github.com/apvee/spfx-react-toolkit)
- [NPM Package](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)
