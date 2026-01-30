import * as React from 'react';
import {
  Stack,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';
import { useSPFxPnP } from '../../../../hooks';
import { InfoRow } from '../shared';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/site-users';

/**
 * Example 2: useSPFxPnP
 * Shows invoke() for single operations and batch() for multiple operations
 */
export const PnPOperationsDemo: React.FC = () => {
  const { invoke, batch, isLoading, error, clearError } = useSPFxPnP();
  const [lists, setLists] = React.useState<Array<{ Title: string; ItemCount: number }>>([]);
  const [batchData, setBatchData] = React.useState<{
    lists: Array<{ Title: string }>;
    user: { Title: string };
    webTitle: string;
  } | null>(null);

  const handleInvokeLists = React.useCallback(async () => {
    try {
      clearError();
      const result = await invoke(sp => 
        sp.web.lists
          .select('Title', 'ItemCount')
          .filter('Hidden eq false')
          .top(10)()
      );
      setLists(result);
    } catch (err) {
      console.error('Invoke error:', err);
    }
  }, [invoke, clearError]);

  const handleBatchOperations = React.useCallback(async () => {
    try {
      clearError();
      // ✅ Correct batch usage: callback receives batchedSP and returns Promise
      const results = await batch(async (batchedSP) => {
        // All these operations will be sent in ONE HTTP request
        const listsPromise = batchedSP.web.lists.select('Title').top(5)();
        const userPromise = batchedSP.web.currentUser.select('Title')();
        const webPromise = batchedSP.web.select('Title')();
        
        // Wait for all batched operations to complete
        return Promise.all([listsPromise, userPromise, webPromise]);
      });

      const [listsResult, userResult, webResult] = results;

      setBatchData({
        lists: listsResult,
        user: userResult,
        webTitle: webResult.Title
      });
    } catch (err) {
      console.error('Batch error:', err);
    }
  }, [batch, clearError]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="CloudUpload" style={{ marginRight: '8px' }} />
        Example 2: useSPFxPnP - Operations & Batching
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      {/* Single Operation Section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Single Operation with invoke()</Label>
        <PrimaryButton
          text={isLoading ? 'Loading...' : 'Load Lists (invoke)'}
          onClick={handleInvokeLists}
          disabled={isLoading}
          iconProps={{ iconName: 'BulletedList' }}
        />
        {lists.length > 0 && (
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { padding: '12px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
            <Label>Lists (top 10, non-hidden):</Label>
            {lists.map((list, idx) => (
              <div key={idx} style={{ padding: '4px 0', borderBottom: idx < lists.length - 1 ? '1px solid #edebe9' : 'none' }}>
                <strong>{list.Title}</strong> - {list.ItemCount} items
              </div>
            ))}
          </Stack>
        )}
      </Stack>

      {/* Batch Operation Section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Batch Operations with batch()</Label>
        <PrimaryButton
          text={isLoading ? 'Loading...' : 'Load Multiple (batch)'}
          onClick={handleBatchOperations}
          disabled={isLoading}
          iconProps={{ iconName: 'Streaming' }}
        />
        {batchData && (
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { padding: '12px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
            <InfoRow label="Web Title" value={batchData.webTitle} icon="CityNext" />
            <InfoRow label="Current User" value={batchData.user.Title} icon="Contact" />
            <Label>Lists (top 5):</Label>
            {batchData.lists.map((list, idx) => (
              <div key={idx} style={{ paddingLeft: '16px' }}>• {list.Title}</div>
            ))}
          </Stack>
        )}
      </Stack>

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        invoke() for single operations, batch() combines multiple requests into ONE HTTP call for better performance.
      </Label>
    </Stack>
  );
};
