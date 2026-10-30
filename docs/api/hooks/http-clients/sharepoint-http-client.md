# useSPFxSPHttpClient

[HTTP client overview](../../INDEX.md#http-client-hooks) · [Usage guide](../../../guides/http-clients.md)

```ts
import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Access SharePoint REST API client.

## Signature

```typescript
function useSPFxSPHttpClient(initialBaseUrl?: string): SPFxSPHttpClientInfo
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `initialBaseUrl` | `string` | No | Base URL for requests (default: current site) |

## Returns

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

## Description

Provides native `SPHttpClient` for SharePoint REST API calls with integrated authentication.

## Example: List Items CRUD

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

## Example: Cross-Site Query

Use the intended URL directly for the request. `setBaseUrl()` updates state for subsequent renders; the current callback still captures the earlier `baseUrl`.

```tsx
import * as React from 'react';
import { useSPFxSPHttpClient } from '@apvee/spfx-react-toolkit';
import { SPHttpClient } from '@microsoft/sp-http';

interface DocumentItem {
  Id: number;
  Title: string;
}

function CrossSiteData() {
  const { invoke, setBaseUrl, isLoading, isReady, error } = useSPFxSPHttpClient();
  const [otherSiteData, setOtherSiteData] = React.useState<DocumentItem[]>([]);

  const loadFromOtherSite = async () => {
    const targetUrl = 'https://tenant.sharepoint.com/sites/OtherSite';
    setBaseUrl(targetUrl);
    try {
      const result = await invoke(client =>
        client.get(
          `${targetUrl}/_api/web/lists/getbytitle('Documents')/items`,
          SPHttpClient.configurations.v1
        ).then(res => res.json())
      );
      setOtherSiteData(result.value);
    } catch {
      // The hook exposes the invocation error for the UI below.
    }
  };

  return <div>
    <button onClick={loadFromOtherSite} disabled={!isReady || isLoading}>
      Load from Other Site
    </button>
    {error && <p>{error.message}</p>}
    {otherSiteData.map(item => <p key={item.Id}>{item.Title}</p>)}
  </div>;
}
```

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxSPHttpClient.ts)

---
