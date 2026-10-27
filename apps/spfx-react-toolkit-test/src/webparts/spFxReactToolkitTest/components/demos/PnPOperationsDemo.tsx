import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Select, Field, Button, Label, MessageBar, MessageBarActions, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
import { InfoRow } from '../shared';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/site-users';
/**
 * Example 2: useSPFxPnP
 * Shows invoke() for single operations and batch() for multiple operations
 */
export const PnPOperationsDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const [operation, setOperation] = React.useState('invoke');
    const mounted = React.useRef(true);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    const activeOperation = React.useRef(operation);
    activeOperation.current = operation;
    const demoStyles = useDemoStyles();
    const { invoke, batch, isLoading, error, clearError } = useSPFxPnP();
    const requestId = React.useRef(0);
    const currentInvoke = React.useRef(invoke);
    currentInvoke.current = invoke;
    const currentBatch = React.useRef(batch);
    currentBatch.current = batch;
    const [lists, setLists] = React.useState<Array<{
        Title: string;
        ItemCount: number;
    }>>([]);
    const [batchData, setBatchData] = React.useState<{
        lists: Array<{
            Title: string;
        }>;
        user: {
            Title: string;
        };
        webTitle: string;
    } | undefined>(undefined);
    React.useLayoutEffect(() => { requestId.current += 1; setLists([]); setBatchData(undefined); }, [invoke, batch]);
    const handleInvokeLists = React.useCallback(async () => {
        const request = ++requestId.current;
        try {
            clearError();
            setLists([]);
            setBatchData(undefined);
            const result = await invoke(sp => sp.web.lists
                .select('Title', 'ItemCount')
                .filter('Hidden eq false')
                .top(10)());
            if (mounted.current && requestId.current === request && currentInvoke.current === invoke && activeOperation.current === 'invoke')
                setLists(result);
        }
        catch (err) {
            console.error('Invoke error:', err);
        }
    }, [invoke, clearError]);
    const handleBatchOperations = React.useCallback(async () => {
        const request = ++requestId.current;
        try {
            clearError();
            setLists([]);
            setBatchData(undefined);
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
            if (mounted.current && requestId.current === request && currentBatch.current === batch && activeOperation.current === 'batch')
                setBatchData({
                    lists: listsResult,
                    user: userResult,
                    webTitle: webResult.Title
                });
        }
        catch (err) {
            console.error('Batch error:', err);
        }
    }, [batch, clearError]);
    return (<DemoCard title="Example 2: useSPFxPnP - Operations & Batching" description="Compare one list request with a batch of web, user and list requests." iconName="CloudUpload" controls={<>{selector?.(false)}<Field label="Request"><Select select={{ 'data-demo-variant': 'PnPOperationsDemo' } as React.SelectHTMLAttributes<HTMLSelectElement>} value={operation} onChange={(_, data) => { requestId.current += 1; setOperation(data.value); setLists([]); setBatchData(undefined); clearError(); }}><option value="invoke">Lists</option><option value="batch">Batch site, user and lists</option></Select></Field></>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={operation === 'invoke' ? handleInvokeLists : handleBatchOperations} disabled={isLoading}>{isLoading ? 'Loading...' : 'Run'}</Button>
  </>} diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      invoke() for single operations, batch() combines multiple requests into ONE HTTP call for better performance.
    </Label>
  </>}>

    {error && (<MessageBar intent="error">
      <MessageBarBody>
        {error.message}
      </MessageBarBody>
      <MessageBarActions>
        <Button aria-label="Dismiss message" onClick={clearError}>Dismiss</Button>
      </MessageBarActions>
    </MessageBar>)}

    {/* Single Operation Section */}

    {operation === 'invoke' && <div className={demoStyles.panel}>
      <Label>Single Operation with invoke()</Label>
      {lists.length > 0 && (<div className={demoStyles.panel}>
        <Label>Lists (top 10, non-hidden):</Label>
        {lists.map((list, idx) => (<div className={demoStyles.observations} key={idx}>
          <strong>{list.Title}</strong> - {list.ItemCount} items
        </div>))}
      </div>)}
    </div>}
    {/* Batch Operation Section */}

    {operation === 'batch' && <div className={demoStyles.panel}>
      <Label>Batch Operations with batch()</Label>
      {batchData && (<div className={demoStyles.panel}>
        <InfoRow label="Web Title" value={batchData.webTitle} icon="CityNext"/>
        <InfoRow label="Current User" value={batchData.user.Title} icon="Contact"/>
        <Label>Lists (top 5):</Label>
        {batchData.lists.map((list, idx) => (<div className={demoStyles.observations} key={idx}>• {list.Title}</div>))}
      </div>)}
    </div>}
  </DemoCard>);
};
