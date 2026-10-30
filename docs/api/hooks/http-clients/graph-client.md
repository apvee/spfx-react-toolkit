# useSPFxMSGraphClient

[HTTP client overview](../../INDEX.md#http-client-hooks) · [Usage guide](../../../guides/http-clients.md)

```ts
import { useSPFxMSGraphClient } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Access Microsoft Graph API client.

## Signature

```typescript
function useSPFxMSGraphClient(): SPFxMSGraphClientInfo
```

## Returns

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

## Description

Provides `MSGraphClientV3` for authenticated Microsoft Graph API access. Requires appropriate Graph permissions granted by admin.

**Async Initialization**: The Graph client requires async initialization via `factory.getClient('3')`. Use `isInitializing`, `initError`, and `isReady` to track initialization state.

## Prerequisites

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

## Type Safety

Install Graph types for TypeScript support:
```bash
npm install @microsoft/microsoft-graph-types --save-dev
```

## Example: User Profile with Full State Handling

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

## Example: Mail Messages

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

## Example: Calendar Events

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

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxMSGraphClient.ts)

---
