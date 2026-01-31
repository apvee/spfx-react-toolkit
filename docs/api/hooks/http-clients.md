# HTTP Client Hooks

> Hooks for accessing SPFx HTTP clients with state management

## Overview

These hooks provide access to SPFx HTTP clients for making API calls with automatic loading and error state management.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxHttpClient`](#usespfxhttpclient) | `SPFxHttpClientInfo` | Generic HTTP client for external APIs |
| [`useSPFxSPHttpClient`](#usespfxsphttpclient) | `SPFxSPHttpClientInfo` | SharePoint REST API client |
| [`useSPFxAadHttpClient`](#usespfxaadhttpclient) | `SPFxAadHttpClientInfo` | Azure AD secured API client |
| [`useSPFxMSGraphClient`](#usespfxmsgraphclient) | `SPFxMSGraphClientInfo` | Microsoft Graph client |

### When to Use Which Client

| Client | Use Case |
|--------|----------|
| `HttpClient` | External APIs, webhooks, public endpoints |
| `SPHttpClient` | SharePoint REST API (`/_api/`) |
| `AadHttpClient` | Custom Azure AD secured APIs |
| `MSGraphClient` | Microsoft Graph API |

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
  /** Native HttpClient from SPFx */
  readonly client: HttpClient;
  
  /** Invoke HTTP call with state management */
  readonly invoke: <T>(fn: (client: HttpClient) => Promise<T>) => Promise<T>;
  
  /** Loading state during invoke() calls */
  readonly isLoading: boolean;
  
  /** Last error from invoke() calls */
  readonly error: Error | undefined;
  
  /** Clear the current error */
  readonly clearError: () => void;
}
```

### Description

Provides native `HttpClient` for generic HTTP requests to external APIs, webhooks, or any non-SharePoint endpoints.

**Two usage patterns:**
1. **invoke()** - Automatic state management (loading + error tracking)
2. **client** - Direct access for full control

### Example: External API

```tsx
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
import { HttpClient } from '@microsoft/sp-http';

function WeatherWidget() {
  const { invoke, isLoading, error, clearError } = useSPFxHttpClient();
  const [weather, setWeather] = React.useState<any>(null);
  
  const loadWeather = () => {
    invoke(client =>
      client.get(
        'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
        HttpClient.configurations.v1
      ).then(res => res.json())
    ).then(data => setWeather(data));
  };
  
  React.useEffect(() => { loadWeather(); }, []);
  
  if (isLoading) return <Spinner />;
  if (error) return <ErrorMessage message={error.message} onRetry={() => { clearError(); loadWeather(); }} />;
  
  return <div>Temperature: {weather?.main?.temp}°K</div>;
}
```

### Example: POST to Webhook

```tsx
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
import { HttpClient } from '@microsoft/sp-http';

function SlackNotifier() {
  const { invoke, isLoading } = useSPFxHttpClient();
  
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
    <button onClick={() => sendNotification('Hello!')} disabled={isLoading}>
      Send to Slack
    </button>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxHttpClient.ts)

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
  /** Native SPHttpClient from SPFx */
  readonly client: SPHttpClient;
  
  /** Current base URL for relative requests */
  readonly baseUrl: string;
  
  /** Set base URL for subsequent requests */
  readonly setBaseUrl: (url: string) => void;
  
  /** GET request with state management */
  readonly get: <T>(endpoint: string) => Promise<T>;
  
  /** POST request with state management */
  readonly post: <T>(endpoint: string, body?: unknown) => Promise<T>;
  
  /** PATCH request with state management */
  readonly patch: <T>(endpoint: string, body?: unknown) => Promise<T>;
  
  /** DELETE request with state management */
  readonly delete: (endpoint: string) => Promise<void>;
  
  /** Invoke custom operation with state management */
  readonly invoke: <T>(fn: (client: SPHttpClient) => Promise<T>) => Promise<T>;
  
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Last error */
  readonly error: Error | undefined;
  
  /** Clear error */
  readonly clearError: () => void;
}
```

### Description

Provides native `SPHttpClient` for SharePoint REST API calls with integrated authentication. Includes convenience methods for common CRUD operations.

### Example: List Items CRUD

```tsx
import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';

interface ITask {
  Id: number;
  Title: string;
  Status: string;
}

function TaskList() {
  const { get, post, patch, delete: del, isLoading, error } = useSPFxSPHttpClient();
  const [tasks, setTasks] = React.useState<ITask[]>([]);
  
  // Load tasks
  const loadTasks = () => {
    get<{ value: ITask[] }>("/_api/web/lists/getbytitle('Tasks')/items")
      .then(result => setTasks(result.value));
  };
  
  // Create task
  const createTask = (title: string) => {
    post("/_api/web/lists/getbytitle('Tasks')/items", {
      Title: title,
      Status: 'New'
    }).then(loadTasks);
  };
  
  // Update task
  const updateTask = (id: number, status: string) => {
    patch(`/_api/web/lists/getbytitle('Tasks')/items(${id})`, {
      Status: status
    }).then(loadTasks);
  };
  
  // Delete task
  const deleteTask = (id: number) => {
    del(`/_api/web/lists/getbytitle('Tasks')/items(${id})`)
      .then(loadTasks);
  };
  
  React.useEffect(() => { loadTasks(); }, []);
  
  if (isLoading) return <Spinner />;
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

function CrossSiteData() {
  const { get, setBaseUrl, isLoading } = useSPFxSPHttpClient();
  const [otherSiteData, setOtherSiteData] = React.useState([]);
  
  const loadFromOtherSite = () => {
    setBaseUrl('https://tenant.sharepoint.com/sites/OtherSite');
    get("/_api/web/lists/getbytitle('Documents')/items")
      .then(result => setOtherSiteData(result.value));
  };
  
  return (
    <button onClick={loadFromOtherSite} disabled={isLoading}>
      Load from Other Site
    </button>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxSPHttpClient.ts)

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
  /** Native AadHttpClient factory from SPFx */
  readonly factory: AadHttpClientFactory;
  
  /** Current AadHttpClient (undefined until getClient called) */
  readonly client: AadHttpClient | undefined;
  
  /** Get client for a specific resource */
  readonly getClient: (resourceUrl: string) => Promise<AadHttpClient>;
  
  /** Invoke API call with state management */
  readonly invoke: <T>(fn: (client: AadHttpClient) => Promise<T>) => Promise<T>;
  
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Last error */
  readonly error: Error | undefined;
  
  /** Clear error */
  readonly clearError: () => void;
}
```

### Description

Provides `AadHttpClient` for calling custom Azure AD secured APIs. Requires API permissions to be configured in SharePoint admin center.

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

### Example: Custom API Call

```tsx
import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
import { AadHttpClient } from '@microsoft/sp-http';

function CustomApiWidget() {
  const { getClient, invoke, isLoading, error } = useSPFxAadHttpClient();
  const [data, setData] = React.useState(null);
  
  const loadData = async () => {
    // Get client for your API's Azure AD app
    await getClient('https://contoso-api.azurewebsites.net');
    
    // Make authenticated request
    invoke(client =>
      client.get(
        'https://contoso-api.azurewebsites.net/api/data',
        AadHttpClient.configurations.v1
      ).then(res => res.json())
    ).then(setData);
  };
  
  React.useEffect(() => { loadData(); }, []);
  
  if (isLoading) return <Spinner />;
  if (error) return <ErrorMessage message={error.message} />;
  
  return <div>{JSON.stringify(data)}</div>;
}
```

### Source

[View source](../../src/hooks/useSPFxAadHttpClient.ts)

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
  /** Native MSGraphClientV3 from SPFx */
  readonly client: MSGraphClientV3 | undefined;
  
  /** Invoke Graph API call with state management */
  readonly invoke: <T>(fn: (client: MSGraphClientV3) => Promise<T>) => Promise<T>;
  
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Last error */
  readonly error: Error | undefined;
  
  /** Clear error */
  readonly clearError: () => void;
}
```

### Description

Provides `MSGraphClientV3` for authenticated Microsoft Graph API access. Requires appropriate Graph permissions granted by admin.

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

### Example: User Profile

```tsx
import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

function UserProfile() {
  const { invoke, isLoading, error, clearError } = useSPFxMSGraphClient();
  const [user, setUser] = React.useState<MicrosoftGraph.User>();
  
  const loadUser = () => {
    invoke(client => 
      client.api('/me')
        .select('displayName,mail,jobTitle,department')
        .get()
    ).then(setUser);
  };
  
  React.useEffect(() => { loadUser(); }, []);
  
  if (isLoading) return <Spinner />;
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
  const { invoke, isLoading } = useSPFxMSGraphClient();
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
  
  React.useEffect(() => { loadMessages(); }, []);
  
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
  const { invoke, isLoading } = useSPFxMSGraphClient();
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
  
  React.useEffect(() => { loadEvents(); }, []);
  
  return (
    <ul>
      {events.map(event => (
        <li key={event.id}>
          {event.subject} - {new Date(event.start?.dateTime).toLocaleString()}
        </li>
      ))}
    </ul>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxMSGraphClient.ts)

---

## See Also

- [PnPjs Hooks](./pnpjs.md) - PnPjs integration for SharePoint
- [Permissions Hooks](./permissions.md) - Permission checking
- [Context Hooks](./context.md) - Context access

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
