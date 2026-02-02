# HTTP Client Hooks

> Hooks for accessing SPFx HTTP clients with state management

## Overview

These hooks provide access to SPFx HTTP clients for making API calls with automatic loading and error state management.

| Hook | Returns | Init Type | Description |
|------|---------|-----------|-------------|
| [`useSPFxHttpClient`](#usespfxhttpclient) | `SPFxHttpClientInfo` | Sync | Generic HTTP client for external APIs |
| [`useSPFxSPHttpClient`](#usespfxsphttpclient) | `SPFxSPHttpClientInfo` | Sync | SharePoint REST API client |
| [`useSPFxAadHttpClient`](#usespfxaadhttpclient) | `SPFxAadHttpClientInfo` | Async | Azure AD secured API client |
| [`useSPFxMSGraphClient`](#usespfxmsgraphclient) | `SPFxMSGraphClientInfo` | Async | Microsoft Graph client |

### When to Use Which Client

| Client | Use Case |
|--------|----------|
| `HttpClient` | External APIs, webhooks, public endpoints |
| `SPHttpClient` | SharePoint REST API (`/_api/`) |
| `AadHttpClient` | Custom Azure AD secured APIs |
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

---

## useSPFxHttpClient

Access generic HTTP client for external API calls.

### Signature

```typescript
function useSPFxHttpClient(): SPFxHttpClientInfo
```

### Returns

```typescript
interface SPFxHttpClientInfo {
  /** Native HttpClient from SPFx (undefined if ServiceScope unavailable) */
  readonly client: HttpClient | undefined;
  
  /** Invoke HTTP call with state management */
  readonly invoke: <T>(fn: (client: HttpClient) => Promise<T>) => Promise<T>;
  
  /** Loading state during invoke() calls */
  readonly isLoading: boolean;
  
  /** Last error from invoke() calls */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** True when client is ready for use */
  readonly isReady: boolean;
}
```

### Description

Provides native `HttpClient` for generic HTTP requests to external APIs, webhooks, or any non-SharePoint endpoints.

**Two usage patterns:**
1. **invoke()** - Automatic state management (loading + error tracking)
2. **client** - Direct access for full control

### Example: External API with isReady

```tsx
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
import { HttpClient } from '@microsoft/sp-http';

function WeatherWidget() {
  const { invoke, isLoading, error, clearError, isReady } = useSPFxHttpClient();
  const [weather, setWeather] = React.useState<any>(null);
  
  const loadWeather = () => {
    invoke(client =>
      client.get(
        'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
        HttpClient.configurations.v1
      ).then(res => res.json())
    ).then(data => setWeather(data));
  };
  
  React.useEffect(() => { 
    if (isReady) loadWeather(); 
  }, [isReady]);
  
  if (!isReady) return <Spinner label="Initializing..." />;
  if (isLoading) return <Spinner label="Loading weather..." />;
  if (error) return <ErrorMessage message={error.message} onRetry={() => { clearError(); loadWeather(); }} />;
  
  return <div>Temperature: {weather?.main?.temp}°K</div>;
}
```

### Example: POST to Webhook

```tsx
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
import { HttpClient } from '@microsoft/sp-http';

function SlackNotifier() {
  const { invoke, isLoading, isReady } = useSPFxHttpClient();
  
  const sendNotification = (message: string) => {
    invoke(client =>
      client.post(
        'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',
        HttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ text: message })
        }
      )
    ).then(() => console.log('Notification sent'));
  };
  
  return (
    <button onClick={() => sendNotification('Hello!')} disabled={!isReady || isLoading}>
      Send to Slack
    </button>
  );
}
```

### Source

[View source](../../../src/hooks/useSPFxHttpClient.ts)

---

## useSPFxSPHttpClient

Access SharePoint REST API client.

### Signature

```typescript
function useSPFxSPHttpClient(initialBaseUrl?: string): SPFxSPHttpClientInfo
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `initialBaseUrl` | `string` | No | Base URL for requests (default: current site) |

### Returns

```typescript
interface SPFxSPHttpClientInfo {
  /** Native SPHttpClient from SPFx (undefined if ServiceScope unavailable) */
  readonly client: SPHttpClient | undefined;
  
  /** Invoke SharePoint REST API call with state management */
  readonly invoke: <T>(fn: (client: SPHttpClient) => Promise<T>) => Promise<T>;
  
  /** Loading state during invoke() calls */
  readonly isLoading: boolean;
  
  /** Last error from invoke() calls */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** Set or change the base URL (for cross-site queries) */
  readonly setBaseUrl: (url: string) => void;
  
  /** Current base URL (site absolute URL) */
  readonly baseUrl: string;
  
  /** True when client is ready for use */
  readonly isReady: boolean;
}
```

### Description

Provides native `SPHttpClient` for SharePoint REST API calls with integrated authentication.

### Example: List Items CRUD

```tsx
import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
import { SPHttpClient } from '@microsoft/sp-http';

interface ITask {
  Id: number;
  Title: string;
  Status: string;
}

function TaskList() {
  const { invoke, isLoading, error, baseUrl, isReady } = useSPFxSPHttpClient();
  const [tasks, setTasks] = React.useState<ITask[]>([]);
  
  // Load tasks
  const loadTasks = () => {
    invoke(client =>
      client.get(
        `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
        SPHttpClient.configurations.v1
      ).then(res => res.json())
    ).then(result => setTasks(result.value));
  };
  
  // Create task
  const createTask = (title: string) => {
    invoke(client =>
      client.post(
        `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items`,
        SPHttpClient.configurations.v1,
        { body: JSON.stringify({ Title: title, Status: 'New' }) }
      ).then(res => res.json())
    ).then(loadTasks);
  };
  
  // Update task
  const updateTask = (id: number, status: string) => {
    invoke(client =>
      client.post(
        `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
          body: JSON.stringify({ Status: status })
        }
      )
    ).then(loadTasks);
  };
  
  // Delete task
  const deleteTask = (id: number) => {
    invoke(client =>
      client.post(
        `${baseUrl}/_api/web/lists/getbytitle('Tasks')/items(${id})`,
        SPHttpClient.configurations.v1,
        { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' } }
      )
    ).then(loadTasks);
  };
  
  React.useEffect(() => { 
    if (isReady) loadTasks(); 
  }, [isReady]);
  
  if (!isReady) return <Spinner label="Initializing..." />;
  if (isLoading) return <Spinner label="Loading tasks..." />;
  if (error) return <ErrorMessage message={error.message} />;
  
  return (
    <ul>
      {tasks.map(task => (
        <li key={task.Id}>
          {task.Title} - {task.Status}
          <button onClick={() => updateTask(task.Id, 'Done')}>Complete</button>
          <button onClick={() => deleteTask(task.Id)}>Delete</button>
        </li>
      ))}
    </ul>
  );
}
```

### Example: Cross-Site Query

```tsx
import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
import { SPHttpClient } from '@microsoft/sp-http';

function CrossSiteData() {
  const { invoke, setBaseUrl, baseUrl, isLoading, isReady } = useSPFxSPHttpClient();
  const [otherSiteData, setOtherSiteData] = React.useState([]);
  
  const loadFromOtherSite = () => {
    setBaseUrl('https://tenant.sharepoint.com/sites/OtherSite');
    invoke(client =>
      client.get(
        `${baseUrl}/_api/web/lists/getbytitle('Documents')/items`,
        SPHttpClient.configurations.v1
      ).then(res => res.json())
    ).then(result => setOtherSiteData(result.value));
  };
  
  return (
    <button onClick={loadFromOtherSite} disabled={!isReady || isLoading}>
      Load from Other Site
    </button>
  );
}
```

### Source

[View source](../../../src/hooks/useSPFxSPHttpClient.ts)

---

## useSPFxAadHttpClient

Access Azure AD secured API client.

### Signature

```typescript
function useSPFxAadHttpClient(initialResourceUrl?: string): SPFxAadHttpClientInfo
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `initialResourceUrl` | `string` | No | Resource URI for the Azure AD app |

### Returns

```typescript
interface SPFxAadHttpClientInfo {
  /** Native AadHttpClient (undefined until initialization completes) */
  readonly client: AadHttpClient | undefined;
  
  /** Invoke API call with state management */
  readonly invoke: <T>(fn: (client: AadHttpClient) => Promise<T>) => Promise<T>;
  
  /** Loading state during invoke() calls */
  readonly isLoading: boolean;
  
  /** Last error from invoke() calls */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** Set or change the resource URL (triggers client re-initialization) */
  readonly setResourceUrl: (url: string) => void;
  
  /** Current Azure AD resource URL or App ID */
  readonly resourceUrl: string | undefined;
  
  /** True while the AAD client is being initialized */
  readonly isInitializing: boolean;
  
  /** Error that occurred during client initialization */
  readonly initError: Error | undefined;
  
  /** True when client is ready for use */
  readonly isReady: boolean;
}
```

### Description

Provides `AadHttpClient` for calling custom Azure AD secured APIs. Requires API permissions to be configured in SharePoint admin center.

**Async Initialization**: Unlike sync hooks, the AAD client requires async initialization via `factory.getClient()`. Use `isInitializing`, `initError`, and `isReady` to track initialization state.

### Prerequisites

1. Register an Azure AD app for your API
2. Request API permissions in `package-solution.json`:
```json
{
  "webApiPermissionRequests": [
    {
      "resource": "contoso-api",
      "scope": "user_impersonation"
    }
  ]
}
```
3. Approve permissions in SharePoint admin center

### Example: Custom API with State Handling

```tsx
import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
import { AadHttpClient } from '@microsoft/sp-http';

function CustomApiWidget() {
  const { 
    invoke, 
    isLoading, 
    error, 
    resourceUrl,
    isInitializing, 
    initError, 
    isReady 
  } = useSPFxAadHttpClient('https://contoso-api.azurewebsites.net');
  
  const [data, setData] = React.useState(null);
  
  const loadData = () => {
    invoke(client =>
      client.get(
        `${resourceUrl}/api/data`,
        AadHttpClient.configurations.v1
      ).then(res => res.json())
    ).then(setData);
  };
  
  React.useEffect(() => {
    if (isReady) loadData();
  }, [isReady]);
  
  // Handle initialization states
  if (isInitializing) return <Spinner label="Initializing AAD client..." />;
  if (initError) return <ErrorMessage message={`Init failed: ${initError.message}`} />;
  if (!isReady) return <Spinner label="Waiting for client..." />;
  
  // Handle operation states
  if (isLoading) return <Spinner label="Loading data..." />;
  if (error) return <ErrorMessage message={error.message} />;
  
  return <div>{JSON.stringify(data)}</div>;
}
```

### Example: Dynamic Resource URL

```tsx
import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
import { AadHttpClient } from '@microsoft/sp-http';

function DynamicApiSelector() {
  const { 
    invoke, 
    setResourceUrl, 
    resourceUrl, 
    isInitializing, 
    isReady,
    clearError 
  } = useSPFxAadHttpClient();
  
  const [data, setData] = React.useState(null);
  
  const loadFromApi = (apiUrl: string) => {
    clearError();
    setResourceUrl(apiUrl);
  };
  
  // Load data when client becomes ready
  React.useEffect(() => {
    if (isReady && resourceUrl) {
      invoke(client =>
        client.get(
          `${resourceUrl}/api/data`,
          AadHttpClient.configurations.v1
        ).then(res => res.json())
      ).then(setData);
    }
  }, [isReady, resourceUrl]);
  
  return (
    <>
      <button onClick={() => loadFromApi('https://api1.contoso.com')} disabled={isInitializing}>
        API 1
      </button>
      <button onClick={() => loadFromApi('https://api2.contoso.com')} disabled={isInitializing}>
        API 2
      </button>
      {isInitializing && <Spinner label="Switching API..." />}
      {data && <pre>{JSON.stringify(data, null, 2)}</pre>}
    </>
  );
}
```

### Source

[View source](../../../src/hooks/useSPFxAadHttpClient.ts)

---

## useSPFxMSGraphClient

Access Microsoft Graph API client.

### Signature

```typescript
function useSPFxMSGraphClient(): SPFxMSGraphClientInfo
```

### Returns

```typescript
interface SPFxMSGraphClientInfo {
  /** Native MSGraphClientV3 (undefined until initialization completes) */
  readonly client: MSGraphClientV3 | undefined;
  
  /** Invoke Graph API call with state management */
  readonly invoke: <T>(fn: (client: MSGraphClientV3) => Promise<T>) => Promise<T>;
  
  /** Loading state during invoke() calls */
  readonly isLoading: boolean;
  
  /** Last error from invoke() calls */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
  
  /** True while the Graph client is being initialized */
  readonly isInitializing: boolean;
  
  /** Error that occurred during client initialization */
  readonly initError: Error | undefined;
  
  /** True when client is ready for use */
  readonly isReady: boolean;
}
```

### Description

Provides `MSGraphClientV3` for authenticated Microsoft Graph API access. Requires appropriate Graph permissions granted by admin.

**Async Initialization**: The Graph client requires async initialization via `factory.getClient('3')`. Use `isInitializing`, `initError`, and `isReady` to track initialization state.

### Prerequisites

1. Request Graph permissions in `package-solution.json`:
```json
{
  "webApiPermissionRequests": [
    {
      "resource": "Microsoft Graph",
      "scope": "User.Read"
    },
    {
      "resource": "Microsoft Graph",
      "scope": "Mail.Read"
    }
  ]
}
```
2. Approve permissions in SharePoint admin center

### Type Safety

Install Graph types for TypeScript support:
```bash
npm install @microsoft/microsoft-graph-types --save-dev
```

### Example: User Profile with Full State Handling

```tsx
import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

function UserProfile() {
  const { 
    invoke, 
    isLoading, 
    error, 
    clearError,
    isInitializing,
    initError,
    isReady 
  } = useSPFxMSGraphClient();
  
  const [user, setUser] = React.useState<MicrosoftGraph.User>();
  
  const loadUser = () => {
    invoke(client => 
      client.api('/me')
        .select('displayName,mail,jobTitle,department')
        .get()
    ).then(setUser);
  };
  
  React.useEffect(() => { 
    if (isReady) loadUser(); 
  }, [isReady]);
  
  // Handle initialization states
  if (isInitializing) return <Spinner label="Initializing Graph client..." />;
  if (initError) return (
    <MessageBar messageBarType={MessageBarType.error}>
      Failed to initialize Graph: {initError.message}
    </MessageBar>
  );
  
  // Handle operation states
  if (isLoading) return <Spinner label="Loading profile..." />;
  if (error) return (
    <MessageBar messageBarType={MessageBarType.error}>
      {error.message}
      <Link onClick={() => { clearError(); loadUser(); }}>Retry</Link>
    </MessageBar>
  );
  
  return (
    <Persona
      text={user?.displayName}
      secondaryText={user?.jobTitle}
      tertiaryText={user?.department}
    />
  );
}
```

### Example: Mail Messages

```tsx
import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

function RecentMail() {
  const { invoke, isLoading, isReady } = useSPFxMSGraphClient();
  const [messages, setMessages] = React.useState<MicrosoftGraph.Message[]>([]);
  
  const loadMessages = () => {
    invoke(client =>
      client.api('/me/messages')
        .version('v1.0')
        .select('subject,from,receivedDateTime')
        .filter("importance eq 'high'")
        .orderBy('receivedDateTime DESC')
        .top(10)
        .get()
    ).then(result => setMessages(result.value));
  };
  
  React.useEffect(() => { 
    if (isReady) loadMessages(); 
  }, [isReady]);
  
  if (!isReady || isLoading) return <Spinner />;
  
  return (
    <List items={messages}>
      {(message) => (
        <ListItem>
          <strong>{message.subject}</strong>
          <span>{message.from?.emailAddress?.name}</span>
        </ListItem>
      )}
    </List>
  );
}
```

### Example: Calendar Events

```tsx
import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

function UpcomingEvents() {
  const { invoke, isLoading, isReady } = useSPFxMSGraphClient();
  const [events, setEvents] = React.useState<MicrosoftGraph.Event[]>([]);
  
  const loadEvents = () => {
    const now = new Date().toISOString();
    const nextWeek = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
    
    invoke(client =>
      client.api('/me/calendarView')
        .query({
          startDateTime: now,
          endDateTime: nextWeek
        })
        .select('subject,start,end,location')
        .orderBy('start/dateTime')
        .top(5)
        .get()
    ).then(result => setEvents(result.value));
  };
  
  React.useEffect(() => { 
    if (isReady) loadEvents(); 
  }, [isReady]);
  
  if (!isReady || isLoading) return <Spinner />;
  
  return (
    <ul>
      {events.map(event => (
        <li key={event.id}>
          {event.subject} - {new Date(event.start?.dateTime || '').toLocaleString()}
        </li>
      ))}
    </ul>
  );
}
```

### Source

[View source](../../../src/hooks/useSPFxMSGraphClient.ts)

---

## See Also

- [OneDrive AppData Hook](./storage.md#usespfxonedriveappdata) - JSON file storage in OneDrive
- [PnPjs Hooks](./pnpjs.md) - PnPjs integration for SharePoint
- [Permissions Hooks](./permissions.md) - Permission checking
- [Context Hooks](./context.md) - Context access

---

*Generated from JSDoc comments. Last updated: February 2, 2026*
