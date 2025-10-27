# SPFx React Toolkit

React runtime & hooks for SharePoint Framework (SPFx). A comprehensive toolkit providing a single `SPFxProvider` with 25+ strongly-typed, enterprise-grade hooks for accessing SPFx context, managing properties, HTTP clients, permissions, storage, performance tracking, and structured logging.

## Features

- 🎯 **Single Provider Setup** - One `SPFxProvider` wrapper for your entire component tree
- 🔄 **Automatic Context Detection** - Detects WebPart, ApplicationCustomizer, CommandSet, or FieldCustomizer
- 💪 **Type-Safe** - Full TypeScript support with strict typing and zero `any` usage
- ⚡ **Optimized Performance** - Jotai-based atomic state management with per-instance scoping
- 🎨 **React Hooks API** - Familiar, ergonomic hooks pattern for all SPFx capabilities
- 📦 **Bidirectional Sync** - Properties automatically sync between UI and Property Pane
- 🔒 **Instance Isolation** - State is scoped per SPFx instance (supports multiple instances)

## Installation

```bash
npm install spfx-react-toolkit jotai
```

## Quick Start

### 1. Wrap Your Component with SPFxProvider

In your WebPart's `render()` method:

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxProvider } from 'spfx-react-toolkit';
import MyComponent from './components/MyComponent';

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  public render(): void {
    const element = React.createElement(
      SPFxProvider,
      { instance: this as never }, // Cast to 'never' to bypass protected properties
      React.createElement(MyComponent)
    );
    
    ReactDom.render(element, this.domElement);
  }
}
```

### 2. Use Hooks in Your Component

```typescript
import * as React from 'react';
import {
  useSPFxProperties,
  useSPFxDisplayMode,
  useSPFxThemeInfo,
  useSPFxUserInfo,
} from 'spfx-react-toolkit';

interface IMyWebPartProps {
  title: string;
  description: string;
}

const MyComponent: React.FC = () => {
  const { properties, setProperties } = useSPFxProperties<IMyWebPartProps>();
  const { isEdit } = useSPFxDisplayMode();
  const theme = useSPFxThemeInfo();
  const { displayName } = useSPFxUserInfo();
  
  return (
    <div style={{ backgroundColor: theme?.semanticColors?.bodyBackground }}>
      <h1>{properties?.title}</h1>
      <p>Welcome, {displayName}!</p>
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

## Available Hooks

### Core & Properties

#### `useSPFxProperties<T>()`
Access and manage SPFx properties with type-safe partial updates and automatic bidirectional sync.

```typescript
const { properties, setProperties, updateProperties } = useSPFxProperties<IMyWebPartProps>();

// Partial update (shallow merge)
setProperties({ title: 'New Title' });

// Updater function pattern
updateProperties(prev => ({ ...prev, title: prev?.title + ' Updated' }));
```

#### `useSPFxDisplayMode()`
Detect edit vs. read mode for conditional rendering.

```typescript
const { mode, isEdit, isRead } = useSPFxDisplayMode();

return isEdit ? <EditPanel /> : <ReadOnlyView />;
```

#### `useSPFxInstanceInfo()`
Get unique instance ID and component kind for debugging and analytics.

```typescript
const { id, kind } = useSPFxInstanceInfo();
// kind: 'WebPart' | 'ApplicationCustomizer' | 'CommandSet' | 'FieldCustomizer'
```

### Context & Environment

#### `useSPFxPageContext()`
Access the full SPFx PageContext object.

```typescript
const pageContext = useSPFxPageContext();
console.log(pageContext.web.absoluteUrl);
```

#### `useSPFxPageType()`
Detect current page type and modern page status.

```typescript
const { pageType, isModernPage, isListView, isSitePage } = useSPFxPageType();
```

#### `useSPFxEnvironmentInfo()`
Detect environment type and local workbench.

```typescript
const { type, isLocal, isClassic, isSharePoint } = useSPFxEnvironmentInfo();

if (isLocal) {
  console.log('Running in local workbench');
}
```

#### `useSPFxTeams()`
Detect Microsoft Teams context and theme.

```typescript
const { supported, theme, teamSiteDomain, teamSitePath } = useSPFxTeams();

if (supported) {
  console.log(`Running in Teams with theme: ${theme}`);
}
```

### User & Site Information

#### `useSPFxUserInfo()`
Access current user information.

```typescript
const { displayName, email, loginName, isAnonymousGuestUser } = useSPFxUserInfo();

return <p>Welcome, {displayName}!</p>;
```

#### `useSPFxSiteInfo()`
Access site metadata and classification.

```typescript
const { title, webUrl, serverRelativeUrl, siteClassification } = useSPFxSiteInfo();
```

#### `useSPFxLocaleInfo()`
Access localization and culture settings.

```typescript
const { uiLocale, currentCultureName } = useSPFxLocaleInfo();
```

#### `useSPFxListInfo()`
Access list context for list-scoped components (returns `undefined` if not in a list).

```typescript
const listInfo = useSPFxListInfo();

if (listInfo) {
  console.log(`List: ${listInfo.title}, ID: ${listInfo.id}`);
}
```

#### `useSPFxHubSiteInfo()`
Access hub site association data.

```typescript
const hubInfo = useSPFxHubSiteInfo();

if (hubInfo?.isHubSite) {
  console.log(`Hub Site URL: ${hubInfo.hubSiteUrl}`);
}
```

### UI & Layout

#### `useSPFxThemeInfo()`
Access theme colors and semantic tokens for styling.

```typescript
const theme = useSPFxThemeInfo();

return (
  <div style={{
    backgroundColor: theme?.semanticColors?.bodyBackground,
    color: theme?.semanticColors?.bodyText
  }}>
    Themed Content
  </div>
);
```

#### `useSPFxContainerSize()`
Get reactive container dimensions that update on resize.

```typescript
const containerSize = useSPFxContainerSize();

if (containerSize && containerSize.width < 600) {
  return <MobileView />;
}
return <DesktopView />;
```

#### `useSPFxContainerInfo()`
Access container DOM element and size tracking status.

```typescript
const { element, size } = useSPFxContainerInfo();

if (size) {
  console.log(`Container: ${size.width}px × ${size.height}px`);
}
```

### Storage

#### `useSPFxSessionStorage<T>()`
Instance-scoped session storage (persists only for current tab/session).

```typescript
const { value: step, setValue: setStep } = useSPFxSessionStorage('wizard-step', 1);

return (
  <div>
    <p>Step: {step}</p>
    <button onClick={() => setStep(s => s + 1)}>Next</button>
  </div>
);
```

#### `useSPFxLocalStorage<T>()`
Instance-scoped local storage (persists across sessions).

```typescript
const { value, setValue } = useSPFxLocalStorage('view-mode', 'grid');

return (
  <div>
    <button onClick={() => setValue('list')}>List View</button>
    <button onClick={() => setValue('grid')}>Grid View</button>
  </div>
);
```

### Performance & Diagnostics

#### `useSPFxPerformance()`
Performance timing API with SPFx context integration.

```typescript
const { mark, measure, time } = useSPFxPerformance();

const fetchData = async () => {
  const result = await time('fetch-data', async () => {
    const response = await fetch('/api/data');
    return response.json();
  });
  
  console.log(`Fetch took ${result.durationMs}ms`);
};
```

#### `useSPFxLogger()`
Structured logging with correlation tracking.

```typescript
const logger = useSPFxLogger();

logger.info('User action', { action: 'click', target: 'button' });
logger.warn('Slow operation detected');
logger.error('Failed to load data', new Error('Network error'));
```

#### `useSPFxCorrelationInfo()`
Access correlation and tenant IDs for tracking.

```typescript
const { correlationId, tenantId } = useSPFxCorrelationInfo();
```

### Permissions & Security

#### `useSPFxPermissions()`
Check user permissions for web and list operations.

```typescript
const { hasWebPermission, hasListPermission } = useSPFxPermissions();

const canManageWeb = hasWebPermission(SPPermission.manageWeb);
const canAddItems = hasListPermission(SPPermission.addListItems, 'list-id');

if (canManageWeb) {
  return <AdminPanel />;
}
```

#### `useSPFxServiceScope()`
Access SPFx service scope for advanced scenarios.

```typescript
const serviceScope = useSPFxServiceScope();

if (serviceScope) {
  // Access SPFx services directly
}
```

### HTTP Clients

#### `useSPFxSPHttpClient()`
Access SharePoint REST API client.

```typescript
const spHttpClient = useSPFxSPHttpClient();

const fetchLists = async () => {
  if (spHttpClient) {
    const response = await spHttpClient.get(
      `${pageContext.web.absoluteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    return data.value;
  }
};
```

#### `useSPFxMSGraphClient()`
Access Microsoft Graph API client.

```typescript
const msGraphClient = useSPFxMSGraphClient();

const fetchMe = async () => {
  if (msGraphClient) {
    const me = await msGraphClient.api('/me').get();
    console.log(me);
  }
};
```

#### `useSPFxAadHttpClient()`
Access Azure AD secured API client.

```typescript
const aadHttpClient = useSPFxAadHttpClient();

const callCustomApi = async () => {
  if (aadHttpClient) {
    const response = await aadHttpClient.get(
      'https://api.contoso.com/data',
      AadHttpClient.configurations.v1
    );
    const data = await response.json();
    return data;
  }
};
```

## Architecture

### State Management
- **Jotai**: Atomic state management with `atomFamily` for per-instance scoping
- **React Context**: Static metadata distribution (instanceId, component type)
- **Automatic Sync**: Bidirectional property sync managed by Provider

### Provider Responsibilities
- Detects SPFx component type (WebPart, Extension, etc.)
- Initializes Jotai atoms scoped to instance ID
- Subscribes to Property Pane changes (WebParts)
- Syncs property updates back to SPFx
- Manages container size tracking
- Handles cleanup on unmount

### Hook Pattern
All hooks follow a consistent pattern:
1. Get instance context via `useSPFxContext()`
2. Access scoped Jotai atoms via `atomFamily(instanceId)`
3. Return read-only interface (no direct atom exposure)
4. Type-safe with full TypeScript inference

## Best Practices

1. **Always wrap with Provider**: Place `SPFxProvider` at the root of your component tree
2. **Type your properties**: Use generics like `useSPFxProperties<IMyProps>()`
3. **Handle undefined**: Some hooks return optional data (list, hub, Teams context)
4. **Use storage wisely**: Session storage for temporary state, local storage for preferences
5. **Leverage performance hooks**: Use `useSPFxPerformance` for critical operations
6. **Log with context**: `useSPFxLogger` includes correlation IDs automatically

## TypeScript Support

All hooks are fully typed with TypeScript. Import types from the library:

```typescript
import type {
  SPFxPropertiesInfo,
  SPFxDisplayModeInfo,
  SPFxThemeInfo,
  SPFxUserInfo,
  SPFxStorageHook,
  // ... and more
} from 'spfx-react-toolkit';
```

## License

MIT

## Contributing

Contributions are welcome! Please open an issue or pull request.

## Support

For issues and questions, please use the [GitHub Issues](https://github.com/apvee/spfx-react-toolkit/issues) page.
