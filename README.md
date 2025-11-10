# SPFx React Toolkit

> A comprehensive React runtime and hooks library for SharePoint Framework (SPFx) with 33+ type-safe hooks. Simplifies SPFx development with instance-scoped state isolation and ergonomic hooks API across WebParts, Extensions, and Command Sets.

![SPFx React Toolkit](./assets/banner.png)

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
  - [WebPart Setup](#webpart-setup)
  - [ApplicationCustomizer Setup](#applicationcustomizer-setup)
  - [FieldCustomizer Setup](#fieldcustomizer-setup)
  - [ListView CommandSet Setup](#listview-commandset-setup)
  - [Using Hooks in Components](#using-hooks-in-components)
- [Core API](#core-api)
  - [Provider Components](#provider-components)
  - [TypeScript Types](#typescript-types)
- [Hooks API](#hooks-api)
  - [Context & Configuration](#context--configuration)
  - [User & Site Information](#user--site-information)
  - [UI & Layout](#ui--layout)
  - [Storage](#storage)
  - [HTTP Clients](#http-clients)
  - [Performance & Diagnostics](#performance--diagnostics)
  - [Permissions & Security](#permissions--security)
  - [Environment & Platform](#environment--platform)
- [PnPjs Integration (Optional)](#pnpjs-integration-optional)
  - [Installation](#pnpjs-installation)
  - [Available PnP Hooks](#available-pnp-hooks)
- [TypeScript Support](#typescript-support)
- [Architecture](#architecture)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Compatibility](#compatibility)
- [License](#license)
- [Contributing](#contributing)
- [Support](#support)

---

## Overview

**SPFx React Toolkit** is a comprehensive React runtime and hooks library for SharePoint Framework (SPFx) development. It provides a single `SPFxProvider` component that wraps your application and enables access to 33+ strongly-typed, production-ready hooks for seamless integration with SPFx context, properties, HTTP clients, permissions, storage, performance tracking, and more.

Built on [Jotai](https://jotai.org/) atomic state management, this toolkit delivers per-instance state isolation, automatic synchronization, and an ergonomic React Hooks API that works across all SPFx component types: **WebParts**, **Application Customizers**, **Field Customizers**, and **ListView Command Sets**.

### Why SPFx React Toolkit?

- **💪 Type-Safe** - Full TypeScript support with zero `any` usage
- **⚡ Optimized** - Jotai atomic state with per-instance scoping
- **🔄 Auto-Sync** - Bidirectional synchronization
- **🎨 Universal** - Works with all SPFx component types
- **📦 Modular** - Tree-shakeable, minimal bundle impact

---

## Features

- ✅ **Automatic Context Detection** - Detects WebPart, ApplicationCustomizer, CommandSet, or FieldCustomizer
- ✅ **33+ React Hooks** - Comprehensive API surface for all SPFx capabilities
- ✅ **Type-Safe** - Full TypeScript inference with strict typing
- ✅ **Instance Isolation** - State scoped per SPFx instance (multi-instance support)
- ✅ **Bidirectional Sync** - Properties automatically sync between UI and SPFx
- ✅ **PnPjs Integration** - Optional hooks for PnPjs v4 with type-safe filters
- ✅ **Performance Tracking** - Built-in hooks for performance measurement and logging
- ✅ **Cross-Platform** - Teams, SharePoint, and Local Workbench support

---

## Installation

```bash
npm install @apvee/spfx-react-toolkit
```

**Auto-Install (npm 7+):**
The following peer dependencies are automatically installed:
- **Jotai** v2+ - Atomic state management (lightweight ~3KB)
- **PnPjs** v4 - SharePoint API operations

### Peer Dependencies

All peer dependencies (`jotai`, `@pnp/sp`, `@pnp/core`, `@pnp/queryable`) are installed automatically with npm 7+. However:

- ✅ **Jotai** (~3KB) - Always included, core dependency for state management
- ✅ **PnP hooks not used?** - Tree-shaking removes unused PnP code (0 KB overhead)
- ✅ **PnP hooks used?** - Only imported parts included (~30-50 KB compressed)
- ✅ **No webpack errors** - All dependencies resolved
- ✅ **No duplicate installations** - npm reuses existing compatible versions

**All hooks available from single import:**

```typescript
import { 
  useSPFxProperties,     // Core hooks
  useSPFxContext,
  useSPFxPnP,           // PnP hooks
  useSPFxPnPList 
} from '@apvee/spfx-react-toolkit';
```

---

## Quick Start

### WebPart Setup

In your WebPart's `render()` method, wrap your component with `SPFxWebPartProvider`:

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxWebPartProvider } from 'spfx-react-toolkit';
import MyComponent from './components/MyComponent';

export interface IMyWebPartProps {
  title: string;
  description: string;
}

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  public render(): void {
    const element = React.createElement(
      SPFxWebPartProvider,
      { instance: this }, // Type-safe, no casting needed
      React.createElement(MyComponent)
    );
    
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

### ApplicationCustomizer Setup

For Application Customizers, use `SPFxApplicationCustomizerProvider`:

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { SPFxApplicationCustomizerProvider } from 'spfx-react-toolkit';
import MyHeaderComponent from './components/MyHeaderComponent';

export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyProps> {
  public onInit(): Promise<void> {
    // Get placeholder
    const placeholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );

    if (placeholder) {
      const element = React.createElement(
        SPFxApplicationCustomizerProvider,
        { instance: this },
        React.createElement(MyHeaderComponent)
      );
      
      ReactDom.render(element, placeholder.domElement);
    }

    return Promise.resolve();
  }
}
```

### FieldCustomizer Setup

For Field Customizers, use `SPFxFieldCustomizerProvider`:

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import { SPFxFieldCustomizerProvider } from 'spfx-react-toolkit';
import MyFieldRenderer from './components/MyFieldRenderer';

export default class MyFieldCustomizer extends BaseFieldCustomizer<IMyProps> {
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const element = React.createElement(
      SPFxFieldCustomizerProvider,
      { instance: this },
      React.createElement(MyFieldRenderer, {
        value: event.fieldValue,
        listItem: event.listItem
      })
    );

    ReactDom.render(element, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDom.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
```

### ListView CommandSet Setup

For ListView Command Sets, use `SPFxListViewCommandSetProvider`:

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPFxListViewCommandSetProvider } from 'spfx-react-toolkit';
import MyDialogComponent from './components/MyDialogComponent';

export default class MyCommandSet extends BaseListViewCommandSet<IMyProps> {
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        this._showDialog(event.selectedRows);
        break;
    }
  }

  private _showDialog(selectedItems: any[]): void {
    const dialogElement = document.createElement('div');
    document.body.appendChild(dialogElement);

    const element = React.createElement(
      SPFxListViewCommandSetProvider,
      { instance: this },
      React.createElement(MyDialogComponent, {
        items: selectedItems,
        onDismiss: () => {
          ReactDom.unmountComponentAtNode(dialogElement);
          document.body.removeChild(dialogElement);
        }
      })
    );

    ReactDom.render(element, dialogElement);
  }
}
```

### Using Hooks in Components

Once wrapped with a Provider, access SPFx capabilities via hooks:

```typescript
import * as React from 'react';
import {
  useSPFxProperties,
  useSPFxDisplayMode,
  useSPFxThemeInfo,
  useSPFxUserInfo,
  useSPFxSiteInfo,
} from 'spfx-react-toolkit';

interface IMyWebPartProps {
  title: string;
  description: string;
}

const MyComponent: React.FC = () => {
  // Access and update properties
  const { properties, setProperties } = useSPFxProperties<IMyWebPartProps>();
  
  // Check display mode
  const { isEdit } = useSPFxDisplayMode();
  
  // Get theme colors
  const theme = useSPFxThemeInfo();
  
  // Get user information
  const { displayName, email } = useSPFxUserInfo();
  
  // Get site information
  const { title: siteTitle, webUrl } = useSPFxSiteInfo();
  
  return (
    <div style={{ 
      backgroundColor: theme?.semanticColors?.bodyBackground,
      color: theme?.semanticColors?.bodyText,
      padding: '20px'
    }}>
      <h1>{properties?.title || 'Default Title'}</h1>
      <p>{properties?.description}</p>
      <p>Welcome, {displayName} ({email})</p>
      <p>Site: {siteTitle} - {webUrl}</p>
      
      {isEdit && (
        <button onClick={() => setProperties({ title: 'Updated Title' })}>
          Update Title
        </button>
      )}
    </div>
  );
};

export default MyComponent;
```

---

## Core API

### Provider Components

The toolkit provides **4 type-safe Provider components**, one for each SPFx component type. Each Provider automatically detects the component kind, initializes instance-scoped state, and enables all hooks.

#### `SPFxWebPartProvider<TProps>`

Type-safe provider for **WebParts**.

**Props:**
- `instance: BaseClientSideWebPart<TProps>` - WebPart instance
- `children?: React.ReactNode` - Child components

**Example:**
```typescript
import { SPFxWebPartProvider } from 'spfx-react-toolkit';

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  public render(): void {
    const element = React.createElement(
      SPFxWebPartProvider,
      { instance: this },
      React.createElement(MyComponent)
    );
    ReactDom.render(element, this.domElement);
  }
}
```

#### `SPFxApplicationCustomizerProvider<TProps>`

Type-safe provider for **Application Customizers**.

**Props:**
- `instance: BaseApplicationCustomizer<TProps>` - Application Customizer instance
- `children?: React.ReactNode` - Child components

**Example:**
```typescript
import { SPFxApplicationCustomizerProvider } from 'spfx-react-toolkit';

export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyProps> {
  public onInit(): Promise<void> {
    const placeholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );

    if (placeholder) {
      const element = React.createElement(
        SPFxApplicationCustomizerProvider,
        { instance: this },
        React.createElement(MyHeaderComponent)
      );
      ReactDom.render(element, placeholder.domElement);
    }

    return Promise.resolve();
  }
}
```

#### `SPFxFieldCustomizerProvider<TProps>`

Type-safe provider for **Field Customizers**.

**Props:**
- `instance: BaseFieldCustomizer<TProps>` - Field Customizer instance
- `children?: React.ReactNode` - Child components

**Example:**
```typescript
import { SPFxFieldCustomizerProvider } from 'spfx-react-toolkit';

export default class MyFieldCustomizer extends BaseFieldCustomizer<IMyProps> {
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const element = React.createElement(
      SPFxFieldCustomizerProvider,
      { instance: this },
      React.createElement(MyFieldRenderer, { value: event.fieldValue })
    );
    ReactDom.render(element, event.domElement);
  }
}
```

#### `SPFxListViewCommandSetProvider<TProps>`

Type-safe provider for **ListView Command Sets**.

**Props:**
- `instance: BaseListViewCommandSet<TProps>` - ListView Command Set instance
- `children?: React.ReactNode` - Child components

**Example:**
```typescript
import { SPFxListViewCommandSetProvider } from 'spfx-react-toolkit';

export default class MyCommandSet extends BaseListViewCommandSet<IMyProps> {
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const element = React.createElement(
      SPFxListViewCommandSetProvider,
      { instance: this },
      React.createElement(MyDialogComponent)
    );
    // Render to dialog container
  }
}
```

### TypeScript Types

All Provider props and hook return types are fully typed and exported:

```typescript
import type {
  // Provider Props
  SPFxWebPartProviderProps,
  SPFxApplicationCustomizerProviderProps,
  SPFxFieldCustomizerProviderProps,
  SPFxListViewCommandSetProviderProps,
  
  // Core Context Types
  HostKind,                    // 'WebPart' | 'AppCustomizer' | 'FieldCustomizer' | 'CommandSet' | 'ACE'
  SPFxComponent,               // Union of all SPFx component types
  SPFxContextType,             // Union of all SPFx context types
  SPFxContextValue,            // Context value: { instanceId, spfxContext, kind }
  ContainerSize,               // { width: number, height: number }
  
  // Hook Return Types
  SPFxPropertiesInfo,          // useSPFxProperties
  SPFxDisplayModeInfo,         // useSPFxDisplayMode
  SPFxInstanceInfo,            // useSPFxInstanceInfo
  SPFxEnvironmentInfo,         // useSPFxEnvironmentInfo
  SPFxPageTypeInfo,            // useSPFxPageType
  SPFxUserInfo,                // useSPFxUserInfo
  SPFxSiteInfo,                // useSPFxSiteInfo
  SPFxLocaleInfo,              // useSPFxLocaleInfo
  SPFxListInfo,                // useSPFxListInfo
  SPFxHubSiteInfo,             // useSPFxHubSiteInfo
  SPFxThemeInfo,               // useSPFxThemeInfo
  SPFxFluent9ThemeInfo,        // useSPFxFluent9ThemeInfo
  SPFxContainerInfo,           // useSPFxContainerInfo
  SPFxStorageHook,             // useSPFxLocalStorage / useSPFxSessionStorage
  SPFxPerformanceInfo,         // useSPFxPerformance
  SPFxPerfResult,              // Performance measurement result
  SPFxLoggerInfo,              // useSPFxLogger
  LogEntry,                    // Log entry structure
  LogLevel,                    // Log levels
  SPFxCorrelationInfo,         // useSPFxCorrelationInfo
  SPFxPermissionsInfo,         // useSPFxPermissions
  SPFxTeamsInfo,               // useSPFxTeams
  TeamsTheme,                  // Teams theme type
  SPFxOneDriveAppDataResult,   // useSPFxOneDriveAppData
  
  // PnP Types (if using PnPjs integration)
  PnPContextInfo,              // useSPFxPnPContext
  SPFxPnPInfo,                 // useSPFxPnP
  SPFxPnPListInfo,             // useSPFxPnPList
  SPFxPnPSearchInfo,           // useSPFxPnPSearch
} from 'spfx-react-toolkit';
```

---

## Hooks API

The toolkit provides **33 specialized hooks** organized by functionality. All hooks are type-safe, memoized, and automatically access the instance-scoped state.

### Context & Configuration

#### `useSPFxPageContext()`

Access the full SharePoint PageContext object containing site, web, user, list, and Teams information.

**Returns:** `PageContext`

**Example:**
```typescript
const pageContext = useSPFxPageContext();

console.log(pageContext.web.title);
console.log(pageContext.web.absoluteUrl);
console.log(pageContext.user.displayName);
```

---

#### `useSPFxProperties<T>()`

Access and manage SPFx properties with type-safe partial updates and automatic bidirectional synchronization with the Property Pane.

**Returns:** `SPFxPropertiesInfo<T>`
- `properties: T | undefined` - Current properties object
- `setProperties: (updates: Partial<T>) => void` - Partial merge update
- `updateProperties: (updater: (current: T | undefined) => T) => void` - Updater function pattern

**Features:**
- ✅ Type-safe with generics
- ✅ Partial updates (shallow merge)
- ✅ Updater function pattern (like React setState)
- ✅ Automatic bidirectional sync with SPFx
- ✅ Property Pane refresh for WebParts

**Example:**
```typescript
interface IMyWebPartProps {
  title: string;
  description: string;
  listId?: string;
}

function MyComponent() {
  const { properties, setProperties, updateProperties } = 
    useSPFxProperties<IMyWebPartProps>();
  
  return (
    <div>
      <h1>{properties?.title ?? 'Default Title'}</h1>
      <p>{properties?.description}</p>
      
      {/* Partial update */}
      <button onClick={() => setProperties({ title: 'New Title' })}>
        Update Title
      </button>
      
      {/* Updater function */}
      <button onClick={() => updateProperties(prev => ({
        ...prev,
        title: (prev?.title ?? '') + ' Updated'
      }))}>
        Append to Title
      </button>
    </div>
  );
}
```

---

#### `useSPFxDisplayMode()`

Access display mode (Read/Edit) for conditional rendering. Display mode is readonly and controlled by SharePoint.

**Returns:** `SPFxDisplayModeInfo`
- `mode: DisplayMode` - Current display mode (Read/Edit)
- `isEdit: boolean` - True if in Edit mode
- `isRead: boolean` - True if in Read mode

**Example:**
```typescript
function MyComponent() {
  const { isEdit, isRead } = useSPFxDisplayMode();
  
  return (
    <div>
      <p>Mode: {isEdit ? 'Editing' : 'Reading'}</p>
      {isEdit && <EditControls />}
      {isRead && <ReadOnlyView />}
    </div>
  );
}
```

---

#### `useSPFxInstanceInfo()`

Get unique instance ID and component kind for debugging, logging, and conditional logic.

**Returns:** `SPFxInstanceInfo`
- `id: string` - Unique identifier for this SPFx instance
- `kind: HostKind` - Component type: `'WebPart' | 'AppCustomizer' | 'FieldCustomizer' | 'CommandSet' | 'ACE'`

**Example:**
```typescript
function MyComponent() {
  const { id, kind } = useSPFxInstanceInfo();
  
  console.log(`Instance ID: ${id}`);
  console.log(`Component Type: ${kind}`);
  
  if (kind === 'WebPart') {
    return <WebPartView />;
  }
  
  return <ExtensionView />;
}
```

---

#### `useSPFxServiceScope()`

Access SPFx ServiceScope for advanced service consumption and dependency injection.

**Returns:** `ServiceScope | undefined`

**Example:**
```typescript
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';

const MyServiceKey = ServiceKey.create<IMyService>('my-service', MyService);

function MyComponent() {
  const serviceScope = useSPFxServiceScope();
  
  const myService = serviceScope?.consume(MyServiceKey);
  
  // Use service...
}
```

---

### User & Site Information

#### `useSPFxUserInfo()`

Access current user information including display name, email, login name, and guest status.

**Returns:** `SPFxUserInfo`
- `displayName: string` - User display name
- `email: string | undefined` - User email address
- `loginName: string` - User login name (e.g., "domain\\user" or email)
- `isExternal: boolean` - Whether user is an external guest

**Example:**
```typescript
function MyComponent() {
  const { displayName, email, isExternal } = useSPFxUserInfo();
  
  return (
    <div>
      <h2>Welcome, {displayName}!</h2>
      {email && <p>Email: {email}</p>}
      {isExternal && <Badge>Guest User</Badge>}
    </div>
  );
}
```

---

#### `useSPFxSiteInfo()`

Access comprehensive site collection and web information with flat, predictable property naming.

**Returns:** `SPFxSiteInfo`

**Web Properties (primary context - 90% use case):**
- `webId: string` - Web ID (GUID)
- `webUrl: string` - Web absolute URL
- `webServerRelativeUrl: string` - Web server-relative URL
- `title: string` - Web display name (most commonly used)
- `languageId: number` - Web language (LCID)
- `logoUrl?: string` - Site logo URL (for branding)

**Site Collection Properties (parent context - specialized):**
- `siteId: string` - Site collection ID (GUID)
- `siteUrl: string` - Site collection absolute URL
- `siteServerRelativeUrl: string` - Site collection server-relative URL
- `siteClassification?: string` - Enterprise classification (e.g., "Confidential", "Public")
- `siteGroup?: SPFxGroupInfo` - Microsoft 365 Group info (if group-connected)

**Example:**
```typescript
function SiteHeader() {
  const { 
    title,              // Web title (most common)
    webUrl,             // Web URL
    logoUrl,            // Site logo
    siteClassification, // Enterprise classification
    siteGroup           // M365 Group info
  } = useSPFxSiteInfo();
  
  return (
    <header>
      {logoUrl && <img src={logoUrl} alt="Site logo" />}
      <h1>{title}</h1>
      <a href={webUrl}>Visit Site</a>
      
      {siteClassification && (
        <Label>Classification: {siteClassification}</Label>
      )}
      
      {siteGroup && (
        <Badge>
          {siteGroup.isPublic ? 'Public Team' : 'Private Team'}
        </Badge>
      )}
    </header>
  );
}
```

---

#### `useSPFxLocaleInfo()`

Access locale and regional settings for internationalization (i18n) with direct Intl API compatibility.

**Returns:** `SPFxLocaleInfo`
- `locale: string` - Current content locale (e.g., "en-US", "it-IT")
- `uiLocale: string` - Current UI language locale
- `timeZone?: SPFxTimeZone` - Time zone information (preview API)
- `isRtl: boolean` - Whether language is right-to-left

**Example:**
```typescript
function DateDisplay() {
  const { locale, isRtl, timeZone } = useSPFxLocaleInfo();
  
  const formatDate = (date: Date) => {
    return new Intl.DateTimeFormat(locale, {
      dateStyle: 'full',
      timeStyle: 'long'
    }).format(date);
  };
  
  return (
    <div dir={isRtl ? 'rtl' : 'ltr'}>
      <p>{formatDate(new Date())}</p>
      {timeZone && (
        <p>Time Zone: {timeZone.description} (UTC {timeZone.offset/60})</p>
      )}
    </div>
  );
}
```

---

#### `useSPFxListInfo()`

Access list information when component is rendered in a list context (Field Customizers, list-scoped components).

**Returns:** `SPFxListInfo | undefined`
- `id: string` - List ID (GUID)
- `title: string` - List title
- `serverRelativeUrl: string` - List server-relative URL
- `baseTemplate?: number` - List template type (e.g., 100 for Generic List, 101 for Document Library)
- `isDocumentLibrary?: boolean` - Whether list is a document library

**Example:**
```typescript
function FieldRenderer() {
  const list = useSPFxListInfo();
  
  if (!list) {
    return <div>Not in list context</div>;
  }
  
  return (
    <div>
      <h3>{list.title}</h3>
      <p>List ID: {list.id}</p>
      {list.isDocumentLibrary && <Icon iconName="DocumentLibrary" />}
    </div>
  );
}
```

---

#### `useSPFxHubSiteInfo()`

Access Hub Site association information with automatic hub URL fetching via REST API.

**Returns:** `SPFxHubSiteInfo`
- `isHubSite: boolean` - Whether site is associated with a hub
- `hubSiteId?: string` - Hub site ID (GUID)
- `hubSiteUrl?: string` - Hub site URL (fetched asynchronously)
- `isLoading: boolean` - Loading state for hub URL fetch
- `error?: Error` - Error during hub URL fetch

**Example:**
```typescript
function HubNavigation() {
  const { isHubSite, hubSiteUrl, isLoading } = useSPFxHubSiteInfo();
  
  if (!isHubSite) return null;
  
  if (isLoading) {
    return <Spinner label="Loading hub info..." />;
  }
  
  return (
    <nav>
      <a href={hubSiteUrl}>← Back to Hub</a>
    </nav>
  );
}
```

---

### UI & Layout

#### `useSPFxThemeInfo()`

Access current SPFx theme (Fluent UI 8) with automatic updates when user switches themes.

**Returns:** `IReadonlyTheme | undefined`

**Example:**
```typescript
function MyComponent() {
  const theme = useSPFxThemeInfo();
  
  return (
    <div style={{
      backgroundColor: theme?.semanticColors?.bodyBackground,
      color: theme?.semanticColors?.bodyText,
      padding: '20px'
    }}>
      Themed Content
    </div>
  );
}
```

---

#### `useSPFxFluent9ThemeInfo()`

Access Fluent UI 9 theme with automatic Teams/SharePoint detection and theme conversion.

**Returns:** `SPFxFluent9ThemeInfo`
- `theme: Theme` - Fluent UI 9 theme object (ready for FluentProvider)
- `isTeams: boolean` - Whether running in Microsoft Teams
- `teamsTheme?: string` - Teams theme name ('default', 'dark', 'contrast')

**Priority order:**
1. Teams native themes (if in Teams)
2. SPFx theme converted to Fluent UI 9
3. Default webLightTheme

**Example:**
```typescript
import { FluentProvider } from '@fluentui/react-components';

function MyWebPart() {
  const { theme, isTeams, teamsTheme } = useSPFxFluent9ThemeInfo();
  
  return (
    <FluentProvider theme={theme}>
      <div>
        <p>Running in: {isTeams ? 'Teams' : 'SharePoint'}</p>
        {isTeams && <p>Teams theme: {teamsTheme}</p>}
        <Button appearance="primary">Themed Button</Button>
      </div>
    </FluentProvider>
  );
}
```

---

#### `useSPFxContainerSize()`

Get reactive container dimensions with Fluent UI 9 aligned breakpoints. Auto-updates on resize.

**Returns:** `SPFxContainerSizeInfo`
- `size: SPFxContainerSize` - Category: 'small' | 'medium' | 'large' | 'xLarge' | 'xxLarge' | 'xxxLarge'
- `isSmall: boolean` - 320-479px (mobile portrait)
- `isMedium: boolean` - 480-639px (mobile landscape)
- `isLarge: boolean` - 640-1023px (tablets)
- `isXLarge: boolean` - 1024-1365px (laptop, desktop)
- `isXXLarge: boolean` - 1366-1919px (wide desktop)
- `isXXXLarge: boolean` - ≥1920px (4K, ultra-wide)
- `width: number` - Actual width in pixels
- `height: number` - Actual height in pixels

**Example:**
```typescript
function ResponsiveWebPart() {
  const { size, isSmall, isXXXLarge, width } = useSPFxContainerSize();
  
  if (isSmall) {
    return <CompactMobileView />;
  }
  
  if (size === 'medium' || size === 'large') {
    return <TabletView />;
  }
  
  if (isXXXLarge) {
    return <UltraWideView columns={6} />; // 4K/ultra-wide
  }
  
  return <DesktopView columns={size === 'xxLarge' ? 4 : 3} />;
}
```

---

#### `useSPFxContainerInfo()`

Access container DOM element and size tracking.

**Returns:** `SPFxContainerInfo`
- `element: HTMLElement | undefined` - Container DOM element
- `size: ContainerSize | undefined` - `{ width: number, height: number }`

**Example:**
```typescript
function MyComponent() {
  const { element, size } = useSPFxContainerInfo();
  
  return (
    <div>
      {size && <p>Container: {size.width}px × {size.height}px</p>}
    </div>
  );
}
```

---

### Storage

#### `useSPFxLocalStorage<T>(key, defaultValue)`

Instance-scoped localStorage for persistent data across sessions. Automatically scoped per SPFx instance.

**Parameters:**
- `key: string` - Storage key (auto-prefixed with instance ID)
- `defaultValue: T` - Default value if not in storage

**Returns:** `SPFxStorageHook<T>`
- `value: T` - Current value
- `setValue: (value: T | ((prev: T) => T)) => void` - Set new value
- `remove: () => void` - Remove value (reset to default)

**Example:**
```typescript
function PreferencesPanel() {
  const { value: viewMode, setValue: setViewMode } = 
    useSPFxLocalStorage('view-mode', 'grid');
  
  return (
    <div>
      <p>View: {viewMode}</p>
      <button onClick={() => setViewMode('list')}>List View</button>
      <button onClick={() => setViewMode('grid')}>Grid View</button>
    </div>
  );
}
```

---

#### `useSPFxSessionStorage<T>(key, defaultValue)`

Instance-scoped sessionStorage for temporary data (current tab/session only). Automatically scoped per SPFx instance.

**Parameters:**
- `key: string` - Storage key (auto-prefixed with instance ID)
- `defaultValue: T` - Default value if not in storage

**Returns:** `SPFxStorageHook<T>`
- `value: T` - Current value
- `setValue: (value: T | ((prev: T) => T)) => void` - Set new value
- `remove: () => void` - Remove value (reset to default)

**Example:**
```typescript
function WizardComponent() {
  const { value: step, setValue: setStep } = 
    useSPFxSessionStorage('wizard-step', 1);
  
  return (
    <div>
      <p>Step: {step} of 5</p>
      <button onClick={() => setStep(s => s + 1)}>Next</button>
      <button onClick={() => setStep(s => s - 1)}>Back</button>
    </div>
  );
}
```

---

### HTTP Clients

#### `useSPFxSPHttpClient()`

Access SharePoint REST API client (SPHttpClient).

**Returns:** `SPFxSPHttpClientInfo`
- `client: SPHttpClient | undefined` - SPHttpClient instance
- `invoke: (fn) => Promise<T>` - Execute with error handling
- `baseUrl: string` - Base URL for REST API calls

**Example:**
```typescript
function ListsViewer() {
  const { invoke, baseUrl } = useSPFxSPHttpClient();
  const [lists, setLists] = useState([]);
  
  const fetchLists = async () => {
    const data = await invoke(async (client) => {
      const response = await client.get(
        `${baseUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return (await response.json()).value;
    });
    setLists(data);
  };
  
  return (
    <div>
      <button onClick={fetchLists}>Load Lists</button>
      <ul>
        {lists.map(list => <li key={list.Id}>{list.Title}</li>)}
      </ul>
    </div>
  );
}
```

---

#### `useSPFxMSGraphClient()`

Access Microsoft Graph API client.

**Returns:** `SPFxMSGraphClientInfo`
- `client: MSGraphClientV3 | undefined` - MS Graph client instance
- `invoke: (fn) => Promise<T>` - Execute with error handling

**Required API Permissions:** Configure in `package-solution.json`

**Example:**
```typescript
function UserProfile() {
  const { invoke } = useSPFxMSGraphClient();
  const [profile, setProfile] = useState(null);
  
  const fetchProfile = async () => {
    const data = await invoke(async (client) => {
      return await client.api('/me').get();
    });
    setProfile(data);
  };
  
  useEffect(() => { fetchProfile(); }, []);
  
  return profile && (
    <div>
      <h3>{profile.displayName}</h3>
      <p>{profile.mail}</p>
    </div>
  );
}
```

---

#### `useSPFxAadHttpClient()`

Access Azure AD secured API client (AadHttpClient).

**Returns:** `SPFxAadHttpClientInfo`
- `client: AadHttpClient | undefined` - AAD HTTP client instance
- `invoke: (fn) => Promise<T>` - Execute with error handling

**Example:**
```typescript
function CustomApiCall() {
  const { invoke } = useSPFxAadHttpClient();
  const [data, setData] = useState(null);
  
  const callApi = async () => {
    const result = await invoke(async (client) => {
      const response = await client.get(
        'https://api.contoso.com/data',
        AadHttpClient.configurations.v1
      );
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return await response.json();
    });
    setData(result);
  };
  
  return (
    <div>
      <button onClick={callApi}>Call API</button>
      {data && <pre>{JSON.stringify(data, null, 2)}</pre>}
    </div>
  );
}
```

---

#### `useSPFxOneDriveAppData<T>(filename, folder?, autoFetch?)`

Manage JSON files in user's OneDrive appRoot folder with unified read/write operations.

**Parameters:**
- `filename: string` - JSON filename
- `folder?: string` - Folder namespace (optional, for isolation)
- `autoFetch?: boolean` - Auto-fetch on mount (default: true)

**Returns:** `SPFxOneDriveAppDataResult<T>`
- `data: T | undefined` - Current data
- `isLoading: boolean` - Loading state
- `isWriting: boolean` - Writing state
- `error?: Error` - Error during operations
- `write: (data: T) => Promise<void>` - Write data to file
- `load: () => Promise<void>` - Manually load data
- `isReady: boolean` - Client ready for operations

**Example:**
```typescript
interface UserSettings {
  theme: string;
  language: string;
}

function SettingsPanel() {
  const { data, write, isLoading, isWriting } = 
    useSPFxOneDriveAppData<UserSettings>('settings.json');
  
  if (isLoading) return <Spinner />;
  
  const handleSave = async () => {
    await write({ theme: 'dark', language: 'en' });
  };
  
  return (
    <div>
      <p>Theme: {data?.theme}</p>
      <button onClick={handleSave} disabled={isWriting}>
        Save Settings
      </button>
    </div>
  );
}
```

---

#### `useSPFxUserPhoto(options?)`

Load user profile photos from Microsoft Graph API. Supports current user or specific users by ID/email.

**Parameters:**
- `options?: UserPhotoOptions` - Optional `{ userId?, email?, size?, autoFetch? }`

**Requires MS Graph Permissions:**
- **User.Read**: For current user's photo
- **User.ReadBasic.All**: For other users' photos

**Returns:** `SPFxUserPhotoInfo`
- `photoUrl: string | undefined` - Photo URL or undefined
- `isLoading: boolean` - Loading state
- `error?: Error` - Error state
- `reload: () => Promise<void>` - Manually reload photo

**Example:**
```typescript
// Current user
const { photoUrl, isLoading } = useSPFxUserPhoto();

// Specific user by email
const { photoUrl } = useSPFxUserPhoto({ 
  email: 'user@contoso.com',
  size: '96x96' // Options: 48x48, 64x64, 96x96, 120x120, 240x240, 360x360, 432x432, 504x504, 648x648
});

// Lazy loading
const { photoUrl, reload, isLoading } = useSPFxUserPhoto({ 
  email: 'user@contoso.com',
  autoFetch: false
});

return (
  <div>
    {photoUrl ? (
      <img src={photoUrl} alt="Avatar" />
    ) : (
      <button onClick={reload} disabled={isLoading}>
        {isLoading ? 'Loading...' : 'Load Photo'}
      </button>
    )}
  </div>
);
```

---

#### `useSPFxTenantProperty<T>(key)`

Manage tenant-wide properties using SharePoint StorageEntity API with smart serialization.

**Parameters:**
- `key: string` - Property key

**Returns:** `SPFxTenantPropertyInfo<T>`
- `data: T | undefined` - Current property value
- `description: string | undefined` - Property description
- `isLoading: boolean` - Loading state
- `error?: Error` - Error state
- `write: (value: T, description?: string) => Promise<void>` - Write property
- `remove: () => Promise<void>` - Remove property
- `canWrite: boolean` - Whether user can write (requires Site Collection Admin on tenant app catalog)

**Requirements:**
- Tenant app catalog must be provisioned
- **Read**: Any authenticated user
- **Write/Remove**: Must be Site Collection Administrator of tenant app catalog

**Smart Serialization:**
- Primitives (string, number, boolean, null, bigint) → stored as string
- Date objects → stored as ISO 8601 string
- Objects/arrays → stored as JSON string

**Example:**
```typescript
// String property
const { data, write, canWrite, isLoading } = useSPFxTenantProperty<string>('appVersion');

if (isLoading) return <Spinner />;

const handleUpdate = async () => {
  if (!canWrite) {
    alert('Insufficient permissions');
    return;
  }
  await write('2.0.1', 'Current application version');
};

return (
  <div>
    <p>Version: {data ?? 'Not Set'}</p>
    {canWrite && <button onClick={handleUpdate}>Update</button>}
  </div>
);

// Number property
const { data: maxSize } = useSPFxTenantProperty<number>('maxUploadSize');
await write(10485760, 'Max file size in bytes');

// Boolean property
const { data: maintenance } = useSPFxTenantProperty<boolean>('maintenanceMode');
if (maintenance) {
  return <MessageBar>System under maintenance</MessageBar>;
}

// Complex object
interface FeatureFlags {
  enableChat: boolean;
  enableAnalytics: boolean;
  maxUsers: number;
}

const { data, write } = useSPFxTenantProperty<FeatureFlags>('featureFlags');

await write({
  enableChat: true,
  enableAnalytics: false,
  maxUsers: 1000
}, 'Global feature flags');

if (data?.enableChat) {
  return <ChatPanel />;
}
```

---

### Performance & Diagnostics

#### `useSPFxPerformance()`

Performance measurement API with automatic SPFx context integration for monitoring and profiling.

**Returns:** `SPFxPerformanceInfo`
- `mark: (name: string) => void` - Create performance mark
- `measure: (name, startMark, endMark?) => SPFxPerfResult` - Measure duration between marks
- `time: <T>(name, fn) => Promise<SPFxPerfResult<T>>` - Time async operations

**Example:**
```typescript
function DataLoader() {
  const { time } = useSPFxPerformance();
  const [data, setData] = useState(null);
  
  const fetchData = async () => {
    const result = await time('fetch-data', async () => {
      const response = await fetch('/api/data');
      return response.json();
    });
    
    console.log(`Fetch took ${result.durationMs}ms`);
    setData(result.result);
  };
  
  return <button onClick={fetchData}>Load Data</button>;
}
```

---

#### `useSPFxLogger(handler?)`

Structured logging with automatic SPFx context (instance ID, user, site, correlation ID).

**Parameters:**
- `handler?: (entry: LogEntry) => void` - Optional custom log handler (e.g., Application Insights)

**Returns:** `SPFxLoggerInfo`
- `debug: (message, extra?) => void` - Log debug message
- `info: (message, extra?) => void` - Log info message
- `warn: (message, extra?) => void` - Log warning message
- `error: (message, extra?) => void` - Log error message

**Example:**
```typescript
function MyComponent() {
  const logger = useSPFxLogger();
  
  const handleClick = () => {
    logger.info('Button clicked', { buttonId: 'save', timestamp: Date.now() });
  };
  
  const handleError = (error: Error) => {
    logger.error('Operation failed', {
      errorMessage: error.message,
      stack: error.stack
    });
  };
  
  return <button onClick={handleClick}>Save</button>;
}
```

---

#### `useSPFxCorrelationInfo()`

Access correlation ID and tenant ID for distributed tracing and diagnostics.

**Returns:** `SPFxCorrelationInfo`
- `correlationId?: string` - Correlation ID for tracking requests
- `tenantId?: string` - Azure AD tenant ID

**Example:**
```typescript
function DiagnosticsPanel() {
  const { correlationId, tenantId } = useSPFxCorrelationInfo();
  
  const logError = (error: Error) => {
    console.error('Error occurred', {
      message: error.message,
      correlationId,
      tenantId,
      timestamp: new Date().toISOString()
    });
  };
  
  return (
    <div>
      <p>Tenant: {tenantId}</p>
      <p>Correlation: {correlationId}</p>
    </div>
  );
}
```

---

### Permissions & Security

#### `useSPFxPermissions()`

Check SharePoint permissions at site, web, and list levels with SPPermission enum helpers.

**Returns:** `SPFxPermissionsInfo`
- `sitePermissions?: SPPermission` - Site collection permissions
- `webPermissions?: SPPermission` - Web permissions
- `listPermissions?: SPPermission` - List permissions (if in list context)
- `hasWebPermission: (permission) => boolean` - Check web permission
- `hasSitePermission: (permission) => boolean` - Check site permission
- `hasListPermission: (permission) => boolean` - Check list permission

**Common Permissions:**
- `SPPermission.manageWeb`
- `SPPermission.addListItems`
- `SPPermission.editListItems`
- `SPPermission.deleteListItems`
- `SPPermission.managePermissions`

**Example:**
```typescript
import { SPPermission } from '@microsoft/sp-page-context';

function AdminPanel() {
  const { hasWebPermission, hasListPermission } = useSPFxPermissions();
  
  const canManage = hasWebPermission(SPPermission.manageWeb);
  const canAddItems = hasListPermission(SPPermission.addListItems);
  
  return (
    <div>
      {canManage && <button>Manage Settings</button>}
      {canAddItems && <button>Add Item</button>}
      {!canManage && <p>Insufficient permissions</p>}
    </div>
  );
}
```

---

#### `useSPFxCrossSitePermissions(siteUrl?, options?)`

Retrieve permissions for a different site/web/list (cross-site permission check).

**Parameters:**
- `siteUrl?: string` - Target site URL (no fetch if undefined/empty - lazy loading)
- `options?: SPFxCrossSitePermissionsOptions` - Optional `{ webUrl?, listId? }`

**Returns:** `SPFxCrossSitePermissionsInfo`
- `sitePermissions?: SPPermission` - Site permissions
- `webPermissions?: SPPermission` - Web permissions
- `listPermissions?: SPPermission` - List permissions (if listId provided)
- `hasWebPermission: (permission) => boolean` - Check web permission
- `hasSitePermission: (permission) => boolean` - Check site permission
- `hasListPermission: (permission) => boolean` - Check list permission
- `isLoading: boolean` - Loading state
- `error?: Error` - Error state

**Example:**
```typescript
import { SPPermission } from '@microsoft/sp-page-context';

function CrossSiteCheck() {
  const [targetUrl, setTargetUrl] = useState<string | undefined>();
  
  const { hasWebPermission, isLoading, error } = useSPFxCrossSitePermissions(
    targetUrl,
    { webUrl: 'https://contoso.sharepoint.com/sites/target/subweb' }
  );
  
  // Trigger fetch
  const checkPermissions = () => {
    setTargetUrl('https://contoso.sharepoint.com/sites/target');
  };
  
  if (isLoading) return <Spinner />;
  if (error) return <MessageBar>{error.message}</MessageBar>;
  
  const canAdd = hasWebPermission(SPPermission.addListItems);
  
  return (
    <div>
      <button onClick={checkPermissions}>Check Permissions</button>
      {targetUrl && <p>Can add items: {canAdd ? 'Yes' : 'No'}</p>}
    </div>
  );
}
```

---

### Environment & Platform

#### `useSPFxEnvironmentInfo()`

Detect execution environment (Local, SharePoint, Teams, Office, Outlook).

**Returns:** `SPFxEnvironmentInfo`
- `type: SPFxEnvironmentType` - 'Local' | 'SharePoint' | 'SharePointOnPrem' | 'Teams' | 'Office' | 'Outlook'
- `isLocal: boolean` - Running in local workbench
- `isWorkbench: boolean` - Running in any workbench
- `isSharePoint: boolean` - SharePoint Online
- `isSharePointOnPrem: boolean` - SharePoint On-Premises
- `isTeams: boolean` - Microsoft Teams
- `isOffice: boolean` - Office application
- `isOutlook: boolean` - Outlook

**Example:**
```typescript
function AdaptiveUI() {
  const { type, isTeams, isLocal } = useSPFxEnvironmentInfo();
  
  if (isLocal) {
    return <DevModeBanner />;
  }
  
  if (isTeams) {
    return <TeamsOptimizedUI />;
  }
  
  return <SharePointUI />;
}
```

---

#### `useSPFxPageType()`

Detect SharePoint page type (modern site page, classic, list page, etc.).

**Returns:** `SPFxPageTypeInfo`
- `pageType: SPFxPageType` - 'sitePage' | 'webPartPage' | 'listPage' | 'listFormPage' | 'profilePage' | 'searchPage' | 'unknown'
- `isModernPage: boolean` - True for modern site pages
- `isSitePage: boolean` - Site page (modern)
- `isListPage: boolean` - List view page
- `isListFormPage: boolean` - List form page
- `isWebPartPage: boolean` - Classic web part page

**Example:**
```typescript
function FeatureGate() {
  const { isModernPage, isSitePage } = useSPFxPageType();
  
  if (!isModernPage) {
    return <div>This feature requires a modern page</div>;
  }
  
  return isSitePage ? <ModernFeature /> : <ClassicFallback />;
}
```

---

#### `useSPFxTeams()`

Access Microsoft Teams context with automatic SDK initialization (v1 and v2 compatible).

**Returns:** `SPFxTeamsInfo`
- `supported: boolean` - Whether Teams context is available
- `context?: unknown` - Teams context object (team, channel, user info)
- `theme?: TeamsTheme` - 'default' | 'dark' | 'highContrast'

**Example:**
```typescript
function TeamsIntegration() {
  const { supported, context, theme } = useSPFxTeams();
  
  if (!supported) {
    return <div>Not running in Teams</div>;
  }
  
  const teamsContext = context as {
    team?: { displayName: string };
    channel?: { displayName: string };
  };
  
  return (
    <div className={`teams-theme-${theme}`}>
      <h3>Team: {teamsContext.team?.displayName}</h3>
      <p>Channel: {teamsContext.channel?.displayName}</p>
    </div>
  );
}
```

---

## PnPjs Integration (Optional)

This toolkit provides optional hooks for working with **PnPjs v4** for SharePoint REST API operations.

### <a id="pnpjs-installation"></a>Installation

All dependencies (Jotai + PnPjs) are **peer dependencies** and installed automatically:

```bash
npm install @apvee/spfx-react-toolkit
# Automatically installs (npm 7+):
# - jotai ^2.0.0
# - @pnp/sp, @pnp/core, @pnp/queryable ^4.0.0
```

**Import Pattern:**

All hooks (core and PnP) are available from the main entry point:

```typescript
import { 
  // Core hooks
  useSPFxProperties,
  useSPFxContext,
  useSPFxSPHttpClient,
  
  // PnP hooks  
  useSPFxPnP,
  useSPFxPnPList,
  useSPFxPnPSearch,
  useSPFxPnPContext 
} from '@apvee/spfx-react-toolkit';
```

**Bundle Size Optimization:**

- **Don't use PnP hooks?** Tree-shaking removes all PnP code from your bundle (0 KB overhead)
- **Use PnP hooks?** Only imported hooks and their dependencies are bundled (~30-50 KB compressed)
- **SPFx bundler optimization:** Webpack automatically excludes unused code

**When to use PnP hooks:**
- ✅ You prefer PnPjs fluent API over native SPHttpClient
- ✅ You need advanced features like batching, caching, selective queries
- ✅ You want cleaner, more maintainable code for SharePoint operations

See [PNPJS_SETUP.md](./PNPJS_SETUP.md) for complete installation and troubleshooting guide.

### Available PnP Hooks

#### `useSPFxPnPContext(siteUrl?, options?)`

Factory hook for creating configured PnPjs SPFI instances with cache, batching, and cross-site support.

**Import:**
```typescript
import { useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
```

**Parameters:
- `siteUrl?: string` - Target site URL (default: current site)
- `options?: PnPContextOptions` - Configuration for cache, batching, etc.

**Returns:** `PnPContextInfo`
- `sp: SPFI | undefined` - Configured SPFI instance
- `isInitialized: boolean` - Whether sp instance is ready
- `error?: Error` - Initialization error
- `siteUrl: string` - Effective site URL

**Example:**
```typescript
// Current site
const { sp, isInitialized, error } = useSPFxPnPContext();

// Cross-site with caching
const hrContext = useSPFxPnPContext('/sites/hr', {
  cache: {
    enabled: true,
    storage: 'session',
    timeout: 600000 // 10 minutes
  }
});
```

---

#### `useSPFxPnP(pnpContext?)`

General-purpose wrapper for any PnP operation with state management and batching.

**Import:**
```typescript
import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
```

**Parameters:
- `pnpContext?: PnPContextInfo` - Optional context from `useSPFxPnPContext` (default: current site)

**Returns:** `SPFxPnPInfo`
- `sp: SPFI | undefined` - SPFI instance for direct access
- `invoke: <T>(fn) => Promise<T>` - Execute single operation with state management
- `batch: <T>(fn) => Promise<T>` - Execute batch operations
- `isLoading: boolean` - Loading state (tracks `invoke`/`batch` calls only)
- `error?: Error` - Error from operations
- `clearError: () => void` - Clear error state
- `isInitialized: boolean` - SP instance ready
- `siteUrl: string` - Effective site URL

**Selective Imports Required:**
```typescript
// Import only what you need
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/search';
```

**Example:**
```typescript
import '@pnp/sp/lists';
import '@pnp/sp/items';

function ListsViewer() {
  const { invoke, batch, isLoading, error } = useSPFxPnP();
  
  // Single operation
  const loadLists = async () => {
    const lists = await invoke(sp => sp.web.lists());
    return lists;
  };
  
  // Batch operation (single HTTP request)
  const loadDashboard = async () => {
    const [user, lists, tasks] = await batch(async (batchedSP) => {
      const user = batchedSP.web.currentUser();
      const lists = batchedSP.web.lists();
      const tasks = batchedSP.web.lists.getByTitle('Tasks').items.top(10)();
      
      return Promise.all([user, lists, tasks]);
    });
    
    return { user, lists, tasks };
  };
  
  if (isLoading) return <Spinner />;
  if (error) return <MessageBar>{error.message}</MessageBar>;
  
  return <button onClick={loadDashboard}>Load Dashboard</button>;
}
```

---

#### `useSPFxPnPList<T>(listTitle, options?, pnpContext?)`

Specialized hook for SharePoint list operations with **type-safe fluent filter API** and CRUD operations.

**Parameters:**
- `listTitle: string` - List title
- `options?: ListQueryOptions` - Query options (filter, select, orderBy, top, etc.)
- `pnpContext?: PnPContextInfo` - Optional context for cross-site operations

**Returns:** `SPFxPnPListInfo<T>`
- `items: T[]` - Current items
- `loading: boolean` - Loading state
- `error?: Error` - Error state
- `hasMore: boolean` - More items available
- `loadMore: () => Promise<void>` - Load next page
- `loadingMore: boolean` - Loading more state
- `refetch: () => Promise<void>` - Reload items
- `create: (data: Partial<T>) => Promise<number>` - Create single item
- `createBatch: (items: Partial<T>[]) => Promise<number[]>` - Batch create
- `update: (id: number, data: Partial<T>) => Promise<void>` - Update item
- `updateBatch: (updates: Array<{id: number, data: Partial<T>}>) => Promise<void>` - Batch update
- `remove: (id: number) => Promise<void>` - Delete item
- `removeBatch: (ids: number[]) => Promise<void>` - Batch delete

**Type-Safe Fluent Filter (PnPjs v4):**
```typescript
interface Task {
  Id: number;
  Title: string;
  Status: string;
  Priority: number;
  DueDate: string;
}

const { items } = useSPFxPnPList<Task>('Tasks', {
  // Type-safe fluent filter (recommended)
  filter: f => f.text("Status").equals("Active")
                .and()
                .number("Priority").greaterThan(3),
  select: ['Id', 'Title', 'Status', 'Priority'],
  orderBy: 'Priority desc',
  top: 50
});
```

**CRUD Operations:**
```typescript
const { items, create, update, remove, createBatch } = useSPFxPnPList<Task>('Tasks');

// Create
const newId = await create({ Title: 'New Task', Status: 'Active' });

// Batch create
const ids = await createBatch([
  { Title: 'Task 1', Status: 'Active' },
  { Title: 'Task 2', Status: 'Active' }
]);

// Update
await update(newId, { Status: 'Completed' });

// Delete
await remove(newId);
```

**Pagination:**
```typescript
const { items, hasMore, loadMore, loadingMore } = useSPFxPnPList<Task>('Tasks', {
  top: 50,
  orderBy: 'Created desc'
});

return (
  <>
    {items.map(item => <TaskCard key={item.Id} task={item} />)}
    {hasMore && (
      <button onClick={loadMore} disabled={loadingMore}>
        {loadingMore ? 'Loading...' : 'Load More'}
      </button>
    )}
  </>
);
```

---

#### `useSPFxPnPSearch(query, options?, pnpContext?)`

Specialized hook for SharePoint Search API with managed properties and refiners.

See [PNPJS_SETUP.md](./PNPJS_SETUP.md) for complete documentation.

---

**For more examples and detailed documentation, see [PNPJS_SETUP.md](./PNPJS_SETUP.md).**

---

## TypeScript Support

All hooks and providers are fully typed with comprehensive TypeScript support. Import types as needed:

```typescript
import type {
  // Core Provider Props
  SPFxWebPartProviderProps,
  SPFxApplicationCustomizerProviderProps,
  SPFxFieldCustomizerProviderProps,
  SPFxListViewCommandSetProviderProps,
  
  // Core Context Types
  HostKind,
  SPFxComponent,
  SPFxContextType,
  SPFxContextValue,
  ContainerSize,
  
  // Hook Return Types
  SPFxPropertiesInfo,
  SPFxDisplayModeInfo,
  SPFxInstanceInfo,
  SPFxEnvironmentInfo,
  SPFxPageTypeInfo,
  SPFxUserInfo,
  SPFxSiteInfo,
  SPFxGroupInfo,
  SPFxLocaleInfo,
  SPFxListInfo,
  SPFxHubSiteInfo,
  SPFxThemeInfo,
  SPFxFluent9ThemeInfo,
  SPFxContainerInfo,
  SPFxContainerSizeInfo,
  SPFxStorageHook,
  SPFxPerformanceInfo,
  SPFxPerfResult,
  SPFxLoggerInfo,
  LogEntry,
  LogLevel,
  SPFxCorrelationInfo,
  SPFxPermissionsInfo,
  SPFxTeamsInfo,
  TeamsTheme,
  
  // HTTP Clients
  SPFxSPHttpClientInfo,
  SPFxMSGraphClientInfo,
  SPFxAadHttpClientInfo,
  SPFxOneDriveAppDataResult,
  
  // PnP Types (if using PnPjs)
  PnPContextInfo,
  SPFxPnPInfo,
  SPFxPnPListInfo,
  SPFxPnPSearchInfo,
} from 'spfx-react-toolkit';
```

**Type Inference:** All hooks provide full type inference when using TypeScript. Use generics where applicable (e.g., `useSPFxProperties<IMyProps>()`) for enhanced type safety.

---

## Architecture

### State Management

The toolkit uses **Jotai** for atomic state management with per-instance scoping:

- **Atomic Design**: Each piece of state (properties, displayMode, theme, etc.) is an independent atom
- **Instance Scoping**: `atomFamily` creates separate atom instances per SPFx component ID
- **Multi-Instance Support**: Multiple WebParts on the same page work independently
- **Minimal Bundle**: Jotai adds only ~3KB to bundle size
- **React-Native**: Built for React, works with Concurrent Mode

### Provider Responsibilities

The `SPFxProvider` (and its type-specific variants) handle:

1. **Component Detection**: Automatically detects WebPart, Extension, or Command Set
2. **Instance Scoping**: Initializes Jotai atoms scoped to unique instance ID
3. **Property Sync**: Subscribes to Property Pane changes (WebParts) and syncs to atoms
4. **Bidirectional Updates**: Syncs hook-based property updates back to SPFx
5. **Container Tracking**: Monitors container size with ResizeObserver
6. **Theme Subscription**: Listens for theme changes and updates atoms
7. **Cleanup**: Proper disposal on unmount

### Hook Pattern

All hooks follow a consistent design:

1. **Access Context**: Get instance metadata via `useSPFxContext()`
2. **Read Atoms**: Access instance-scoped atoms via `atomFamily(instanceId)`
3. **Return Interface**: Provide read-only or controlled interfaces (no direct atom exposure)
4. **Type Safety**: Full TypeScript inference with zero `any` usage

### Why Jotai?

- ✅ **Atomic**: Independent state units prevent unnecessary re-renders
- ✅ **Scoped**: `atomFamily` enables perfect isolation between instances
- ✅ **Minimal**: Small bundle size (~3KB)
- ✅ **Modern**: Built for React, supports Concurrent Mode
- ✅ **TypeScript-First**: Excellent type inference

---

## Best Practices

1. **Always Use Provider**: Wrap your entire component tree with the appropriate Provider
2. **Type Your Properties**: Use generics like `useSPFxProperties<IMyProps>()` for type safety
3. **Handle Undefined**: Some hooks return optional data (list info, hub info, Teams context)
4. **Storage Keys**: Use descriptive keys for localStorage/sessionStorage
5. **Performance Monitoring**: Leverage `useSPFxPerformance` for critical operations
6. **Structured Logging**: Use `useSPFxLogger` with correlation IDs for better diagnostics
7. **Error Handling**: Always wrap HTTP client calls in try-catch blocks
8. **Memoization**: Use `useMemo`/`useCallback` for expensive computations based on hook data
9. **Responsive Design**: Use `useSPFxContainerSize` for adaptive layouts
10. **Permission Checks**: Gate features with `useSPFxPermissions` for better UX

---

## Troubleshooting

### Provider Errors

**Error: "useSPFxContext must be used within SPFxProvider"**
- Ensure your component tree is wrapped with the appropriate Provider
- Verify hooks are called inside functional components, not outside
- Check that the Provider is mounted before hook calls

### Property Sync Issues

- **Properties not updating?** Use `setProperties` from `useSPFxProperties`, not direct SPFx property mutation
- **Property Pane not reflecting changes?** Ensure SPFx instance properties are mutable
- **Sync delays?** Property sync is intentional - hooks update immediately, Property Pane follows

### Teams Context Not Available

- Teams context loads asynchronously - always check `supported` flag before using
- In local workbench, Teams context won't be available
- Requires SPFx to be running inside Teams environment

### Storage Not Persisting

- Check browser settings - localStorage/sessionStorage may be disabled
- Storage keys are automatically scoped per instance - different instances have isolated storage
- Session storage clears when tab closes (by design)

### Type Errors

- Import types from the library: `import type { ... } from 'spfx-react-toolkit'`
- Use type assertions when accessing SPFx context-specific properties
- Check that generics are properly specified (e.g., `useSPFxProperties<IMyProps>()`)

---

## Compatibility

- **SPFx Version**: >=1.18.0
- **Node.js**: Node.js version aligned with your SPFx version (e.g., Node 18.x for SPFx 1.21.1 - see [SPFx compatibility table](https://learn.microsoft.com/sharepoint/dev/spfx/compatibility))
- **React**: 17.x (SPFx standard)
- **TypeScript**: ~5.3.3
- **Jotai**: ^2.0.0
- **Browsers**: Modern browsers (Chrome, Edge, Firefox, Safari)
- **SharePoint**: SharePoint Online
- **Microsoft 365**: Teams, Office, Outlook (with SPFx support)

---

## License

MIT License

---

## Links

- **GitHub Repository**: [https://github.com/apvee/spfx-react-toolkit](https://github.com/apvee/spfx-react-toolkit)
- **Issues**: [https://github.com/apvee/spfx-react-toolkit/issues](https://github.com/apvee/spfx-react-toolkit/issues)
- **NPM Package**: [@apvee/spfx-react-toolkit](https://www.npmjs.com/package/@apvee/spfx-react-toolkit)

---

Made with ❤️ by [Apvee Solutions](https://github.com/apvee)