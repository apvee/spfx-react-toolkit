import * as React from 'react';
import {
  Stack,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Icon,
  Label,
} from '@fluentui/react';
import { useSPFxHttpClient } from '../../../../hooks';
import { HttpClient } from '@microsoft/sp-http';
import { InfoRow } from '../shared';

/**
 * Example: useSPFxHttpClient
 * Demonstrates calling external public APIs using HttpClient
 */
export const HttpClientDemo: React.FC = () => {
  const { invoke, isLoading, error, clearError } = useSPFxHttpClient();
  const [todos, setTodos] = React.useState<Array<{ id: number; title: string; completed: boolean }>>([]);
  const [selectedTodo, setSelectedTodo] = React.useState<{ id: number; title: string; completed: boolean; userId: number } | null>(null);

  const loadTodos = React.useCallback(async () => {
    try {
      const data = await invoke(client =>
        client.get(
          'https://jsonplaceholder.typicode.com/todos?_limit=5',
          HttpClient.configurations.v1
        ).then(res => res.json())
      );
      setTodos(data);
      setSelectedTodo(null);
    } catch (err) {
      console.error('Failed to load todos:', err);
    }
  }, [invoke]);

  const loadTodoDetails = React.useCallback(async (id: number) => {
    try {
      const data = await invoke(client =>
        client.get(
          `https://jsonplaceholder.typicode.com/todos/${id}`,
          HttpClient.configurations.v1
        ).then(res => res.json())
      );
      setSelectedTodo(data);
    } catch (err) {
      console.error('Failed to load todo details:', err);
    }
  }, [invoke]);

  return (
    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { padding: '16px', border: '1px solid #ddd', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="CloudDownload" style={{ marginRight: '8px' }} />
        HttpClient Example - Public API Call
      </h3>
      <MessageBar messageBarType={MessageBarType.info}>
        This example demonstrates calling a public REST API (JSONPlaceholder) using <strong>useSPFxHttpClient</strong>.
        The hook provides automatic state management (loading/error) for external HTTP calls.
      </MessageBar>

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton
          text="Load Todos"
          onClick={loadTodos}
          disabled={isLoading}
          iconProps={{ iconName: 'Download' }}
        />
        {error && (
          <DefaultButton
            text="Clear Error"
            onClick={clearError}
            iconProps={{ iconName: 'Clear' }}
          />
        )}
      </Stack>

      {isLoading && (
        <MessageBar messageBarType={MessageBarType.info}>
          <Icon iconName="Sync" style={{ marginRight: '8px' }} />
          Loading data from external API...
        </MessageBar>
      )}

      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          <strong>Error:</strong> {error.message}
        </MessageBar>
      )}

      {todos.length > 0 && (
        <Stack tokens={{ childrenGap: 10 }}>
          <Label>Todos from JSONPlaceholder API:</Label>
          {todos.map(todo => (
            <Stack
              key={todo.id}
              horizontal
              tokens={{ childrenGap: 10 }}
              styles={{
                root: {
                  padding: '8px',
                  border: '1px solid #eee',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  ':hover': { backgroundColor: '#f5f5f5' }
                }
              }}
              onClick={() => loadTodoDetails(todo.id)}
            >
              <Icon iconName={todo.completed ? 'CompletedSolid' : 'CircleRing'} />
              <span>{todo.title}</span>
            </Stack>
          ))}
        </Stack>
      )}

      {selectedTodo && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { padding: '12px', backgroundColor: '#f0f0f0', borderRadius: '4px' } }}>
          <Label>Selected Todo Details:</Label>
          <InfoRow label="ID" value={String(selectedTodo.id)} icon="NumberField" />
          <InfoRow label="User ID" value={String(selectedTodo.userId)} icon="Contact" />
          <InfoRow label="Title" value={selectedTodo.title} icon="TextDocument" />
          <InfoRow label="Completed" value={selectedTodo.completed ? 'Yes' : 'No'} icon={selectedTodo.completed ? 'CompletedSolid' : 'CircleRing'} />
        </Stack>
      )}
    </Stack>
  );
};
