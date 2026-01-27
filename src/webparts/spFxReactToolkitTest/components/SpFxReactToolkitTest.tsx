import * as React from 'react';
import styles from './SpFxReactToolkitTest.module.scss';
import {
  Stack,
  Pivot,
  PivotItem,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';

/* eslint-disable max-lines */

// Import all hooks from SPFx React Toolkit
import {
  useSPFxProperties,
  useSPFxThemeInfo,
  useSPFxUserInfo,
  useSPFxEnvironmentInfo,
  useSPFxSiteInfo,
  useSPFxDisplayMode,
  useSPFxInstanceInfo,
  useSPFxPageContext,
  useSPFxTeams,
  useSPFxListInfo,
  useSPFxLocaleInfo,
  useSPFxHubSiteInfo,
  useSPFxCorrelationInfo,
  useSPFxPermissions,
  useSPFxCrossSitePermissions,
  useSPFxContainerSize,
  useSPFxContainerInfo,
  useSPFxSessionStorage,
  useSPFxLocalStorage,
  useSPFxLogger,
  useSPFxPageType,
  useSPFxServiceScope,
  useSPFxSPHttpClient,
  useSPFxMSGraphClient,
  useSPFxAadHttpClient,
  useSPFxHttpClient,
  useSPFxPerformance,
  useSPFxFluent9ThemeInfo,
  useSPFxOneDriveAppData,
  useSPFxTenantProperty,
  useSPFxUserPhoto,
  useSPFxPnPContext,
  useSPFxPnP,
  useSPFxPnPList,
  useSPFxPnPSearch,
  SearchVerticals,
} from '../../../hooks';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';
import { SPPermission } from '@microsoft/sp-page-context';
import { HttpClient } from '@microsoft/sp-http';

interface IWebPartProps {
  description: string;
}

interface IOneDriveTestData {
  message: string;
  counter: number;
  timestamp: number;
}

// =============================================
// Helper Components (Must be defined first)
// =============================================

// Helper component for displaying info rows
const InfoRow: React.FC<{ label: string; value: string | undefined; icon?: string }> = ({ label, value, icon }) => (
  <div className={styles.infoRow}>
    {icon && <Icon iconName={icon} />}
    <Label>{label}:</Label>
    <span>{value || 'N/A'}</span>
  </div>
);

// Helper component for status badges
const StatusBadge: React.FC<{ available: boolean; label: string }> = ({ available, label }) => (
  <div className={`${styles.statusBadge} ${available ? styles.success : styles.error}`}>
    <Icon iconName={available ? 'Completed' : 'ErrorBadge'} />
    {label}
  </div>
);

// =============================================
// HttpClient Example Component
// =============================================

/**
 * Example: useSPFxHttpClient
 * Demonstrates calling external public APIs using HttpClient
 */
const HttpClientExample: React.FC = () => {
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

// =============================================
// PnPjs Example Components
// =============================================

/**
 * Example 1: useSPFxPnPContext
 * Shows how to create PnP contexts for current and cross-site scenarios
 */
const PnPContextExample: React.FC = () => {
  const [crossSiteUrl, setCrossSiteUrl] = React.useState('');
  const [siteInfo, setSiteInfo] = React.useState<{ title: string; url: string; description: string } | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | null>(null);

  // Current site context (default)
  const currentContext = useSPFxPnPContext();

  // Cross-site context (only when URL is provided)
  const crossSiteContext = useSPFxPnPContext(crossSiteUrl || undefined);

  const handleLoadCurrentSite = React.useCallback(async () => {
    if (!currentContext.sp) return;

    setLoading(true);
    setErrorMsg(null);
    try {
      const web = await currentContext.sp.web.select('Title', 'Url', 'Description')();
      setSiteInfo({
        title: web.Title,
        url: web.Url,
        description: web.Description || '(no description)'
      });
    } catch (error) {
      console.error('Error loading site:', error);
      setErrorMsg(error instanceof Error ? error.message : 'Unknown error');
      setSiteInfo(null);
    } finally {
      setLoading(false);
    }
  }, [currentContext.sp]);

  const handleLoadCrossSite = React.useCallback(async () => {
    if (!crossSiteContext.sp || !crossSiteUrl) return;

    setLoading(true);
    setErrorMsg(null);
    try {
      const web = await crossSiteContext.sp.web.select('Title', 'Url', 'Description')();
      setSiteInfo({
        title: web.Title,
        url: web.Url,
        description: web.Description || '(no description)'
      });
    } catch (error) {
      console.error('Error loading cross-site:', error);
      setErrorMsg(error instanceof Error ? error.message : 'Unknown error');
      setSiteInfo(null);
    } finally {
      setLoading(false);
    }
  }, [crossSiteContext.sp, crossSiteUrl]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="Globe" style={{ marginRight: '8px' }} />
        Example 1: useSPFxPnPContext - Site Information
      </h3>
      <Separator />

      {errorMsg && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg(null)}>
          {errorMsg}
        </MessageBar>
      )}

      {/* Current Site Section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Current Site Context</Label>
        <InfoRow label="Effective URL" value={currentContext.siteUrl} icon="Link" />
        <InfoRow label="Is Initialized" value={currentContext.isInitialized ? 'Yes' : 'No'} icon="CheckMark" />
        {currentContext.error && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Init Error: {currentContext.error.message}
          </MessageBar>
        )}
        <PrimaryButton
          text={loading ? 'Loading...' : 'Load Current Site Info'}
          onClick={handleLoadCurrentSite}
          disabled={!currentContext.isInitialized || loading}
          iconProps={{ iconName: 'CloudDownload' }}
        />
      </Stack>

      {/* Cross-Site Section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Cross-Site Context (Optional)</Label>
        <TextField
          label="Site URL"
          value={crossSiteUrl}
          onChange={(_, newValue) => setCrossSiteUrl(newValue ?? '')}
          placeholder="e.g., /sites/hr or https://contoso.sharepoint.com/sites/hr"
          description="Leave empty to use current site"
        />
        {crossSiteUrl && (
          <>
            <InfoRow label="Resolved URL" value={crossSiteContext.siteUrl} icon="Link" />
            <InfoRow label="Is Initialized" value={crossSiteContext.isInitialized ? 'Yes' : 'No'} icon="CheckMark" />
            {crossSiteContext.error && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Init Error: {crossSiteContext.error.message}
              </MessageBar>
            )}
            <PrimaryButton
              text={loading ? 'Loading...' : 'Load Cross-Site Info'}
              onClick={handleLoadCrossSite}
              disabled={!crossSiteContext.isInitialized || loading}
              iconProps={{ iconName: 'Globe' }}
            />
          </>
        )}
      </Stack>

      {/* Results */}
      {siteInfo && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { padding: '12px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
          <Label style={{ fontWeight: 600 }}>
            <Icon iconName="CheckMark" style={{ marginRight: '4px', color: '#107c10' }} />
            Site Information Loaded:
          </Label>
          <InfoRow label="Title" value={siteInfo.title} icon="CityNext" />
          <InfoRow label="URL" value={siteInfo.url} icon="Link" />
          <InfoRow label="Description" value={siteInfo.description} icon="Info" />
        </Stack>
      )}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        useSPFxPnPContext creates configured SPFI instances. Use for cross-site scenarios or when you need custom cache/batch config.
      </Label>
    </Stack>
  );
};

/**
 * Example 2: useSPFxPnP
 * Shows invoke() for single operations and batch() for multiple operations
 */
const PnPOperationsExample: React.FC = () => {
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

/**
 * Example 3: useSPFxPnPSearch - Basic Search
 */
const PnPSearchBasicExample: React.FC = () => {
  const [searchText, setSearchText] = React.useState('');
  
  const {
    search,
    results,
    totalResults,
    loading,
    error,
    clearError,
  } = useSPFxPnPSearch<{ Title: string; Path: string; FileType: string }>({
    pageSize: 10
  });

  const handleSearch = React.useCallback(async () => {
    if (!searchText.trim()) return;
    try {
      const searchResults = await search(searchText);
      console.log('[PnPSearchBasic] Search completed:', {
        returnedResults: searchResults.length,
        stateResults: results.length,
        totalResults
      });
    } catch (err) {
      console.error('Search error:', err);
    }
  }, [searchText, search, results.length, totalResults]);

  // Debug: log quando results cambia
  React.useEffect(() => {
    console.log('[PnPSearchBasic] Results state updated:', {
      resultsLength: results.length,
      totalResults,
      loading
    });
  }, [results, totalResults, loading]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="Search" style={{ marginRight: '8px' }} />
        Example 1: Basic Search
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <TextField
          value={searchText}
          onChange={(_, newValue) => setSearchText(newValue ?? '')}
          placeholder="Search SharePoint..."
          styles={{ root: { flexGrow: 1 } }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              handleSearch().catch(err => console.error('Search error:', err));
            }
          }}
        />
        <PrimaryButton
          text={loading ? 'Searching...' : 'Search'}
          onClick={handleSearch}
          disabled={!searchText.trim() || loading}
          iconProps={{ iconName: 'Search' }}
        />
      </Stack>

      {totalResults !== undefined && (
        <MessageBar messageBarType={MessageBarType.info}>
          Found {totalResults} result{totalResults !== 1 ? 's' : ''}
        </MessageBar>
      )}

      {results.length > 0 ? (
        <Stack tokens={{ childrenGap: 6 }}>
          {results.map((result) => (
            <Stack
              key={result.id}
              tokens={{ childrenGap: 4 }}
              styles={{ root: { padding: '10px', backgroundColor: '#faf9f8', borderRadius: '4px' } }}
            >
              <div style={{ fontWeight: 600 }}>{result.data.Title || '(No Title)'}</div>
              <a href={result.data.Path} target="_blank" rel="noopener noreferrer" style={{ fontSize: '12px' }}>
                {result.data.Path}
              </a>
              {result.data.FileType && <Label>Type: {result.data.FileType}</Label>}
            </Stack>
          ))}
        </Stack>
      ) : totalResults === 0 ? (
        <MessageBar>No results found. Try a different query.</MessageBar>
      ) : null}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Simple text search. Try: &quot;document&quot;, &quot;report&quot;, &quot;ContentType:Document&quot;
      </Label>
    </Stack>
  );
};

/**
 * Example 4: Advanced Search with Builder and Verticals
 */
const PnPSearchAdvancedExample: React.FC = () => {
  const [selectedVertical, setSelectedVertical] = React.useState<string>('All');
  
  const {
    search,
    results,
    totalResults,
    loading,
    loadMore,
    hasMore,
    error,
    clearError,
  } = useSPFxPnPSearch<{ 
    Title: string; 
    Path: string; 
    Author: string; 
    LastModifiedTime: string;
    FileType: string;
  }>({
    pageSize: 5,
    selectProperties: ['Title', 'Path', 'Author', 'LastModifiedTime', 'FileType']
  });

  const getVerticalSourceId = (vertical: string): string | undefined => {
    switch (vertical) {
      case 'People': return SearchVerticals.People;
      case 'Documents': return SearchVerticals.Documents;
      case 'Pages': return SearchVerticals.Pages;
      case 'Videos': return SearchVerticals.Videos;
      default: return undefined;
    }
  };

  const handleSearch = React.useCallback((query: string) => {
    search((builder) => {
      let result = builder.text(query).sortList({ Property: 'LastModifiedTime', Direction: 1 });
      const sourceId = getVerticalSourceId(selectedVertical);
      if (sourceId) result = result.sourceId(sourceId);
      return result;
    }).catch(err => console.error('Search error:', err));
  }, [search, selectedVertical]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="SearchAndApps" style={{ marginRight: '8px' }} />
        Example 2: Advanced Search with Builder
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
        {['All', 'Documents', 'Pages', 'People', 'Videos'].map((v) => (
          <DefaultButton
            key={v}
            text={v}
            onClick={() => setSelectedVertical(v)}
            primary={selectedVertical === v}
          />
        ))}
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 6 }}>
        <PrimaryButton
          text="Search All"
          onClick={() => handleSearch('*')}
          disabled={loading}
        />
        <DefaultButton
          text="Search Documents"
          onClick={() => handleSearch('filetype:docx OR filetype:pdf')}
          disabled={loading}
        />
      </Stack>

      {totalResults !== undefined && (
        <MessageBar messageBarType={MessageBarType.info}>
          Found {totalResults} result{totalResults !== 1 ? 's' : ''} in &quot;{selectedVertical}&quot;
        </MessageBar>
      )}

      {results.length > 0 && (
        <Stack tokens={{ childrenGap: 6 }}>
          {results.map((result) => (
            <Stack key={result.id} tokens={{ childrenGap: 2 }} styles={{ root: { padding: '10px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
              <div style={{ fontWeight: 600 }}>{result.data.Title || '(No Title)'}</div>
              <Label>{result.data.Author} - {result.data.LastModifiedTime ? new Date(result.data.LastModifiedTime).toLocaleDateString() : 'N/A'}</Label>
              <a href={result.data.Path} target="_blank" rel="noopener noreferrer" style={{ fontSize: '11px' }}>
                {result.data.Path}
              </a>
            </Stack>
          ))}
          {hasMore && (
            <PrimaryButton text={loading ? 'Loading...' : 'Load More'} onClick={loadMore} disabled={loading} />
          )}
        </Stack>
      )}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Builder API with verticals (People, Documents, Pages, Videos) and pagination.
      </Label>
    </Stack>
  );
};

/**
 * Example 5: Faceted Search with Refiners
 */
const PnPSearchRefinersExample: React.FC = () => {
  const {
    search,
    results,
    refiners,
    totalResults,
    loading,
    applyRefiner,
    error,
    clearError,
  } = useSPFxPnPSearch<{ Title: string; Path: string; FileType: string; Author: string }>({
    pageSize: 10,
    selectProperties: ['Title', 'Path', 'FileType', 'Author'],
    refiners: 'FileType,Author'
  });

  const handleSearch = React.useCallback(() => {
    search('*').catch(err => console.error('Search error:', err));
  }, [search]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="FilterSettings" style={{ marginRight: '8px' }} />
        Example 3: Faceted Search with Refiners
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      <PrimaryButton
        text={loading ? 'Loading...' : 'Search All Content'}
        onClick={handleSearch}
        disabled={loading}
        iconProps={{ iconName: 'Search' }}
      />

      {totalResults !== undefined && (
        <Label>
          <Icon iconName="SearchIssue" style={{ marginRight: '4px', color: '#0078d4' }} />
          Found {totalResults} results
        </Label>
      )}

      <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { alignItems: 'flex-start' } }}>
        {refiners.length > 0 && (
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: '200px', padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
            <Label style={{ fontWeight: 600 }}>Filters</Label>
            {refiners.map((refiner) => (
              <Stack key={refiner.name} tokens={{ childrenGap: 2 }}>
                <Label>{refiner.name}</Label>
                {refiner.entries.slice(0, 5).map((entry) => (
                  <DefaultButton
                    key={entry.value}
                    text={`${entry.value} (${entry.count})`}
                    onClick={() => applyRefiner(refiner.name, entry.value)}
                    styles={{ root: { height: '28px', fontSize: '12px' } }}
                  />
                ))}
              </Stack>
            ))}
          </Stack>
        )}

        <Stack tokens={{ childrenGap: 6 }} styles={{ root: { flexGrow: 1 } }}>
          {results.length > 0 ? (
            results.map((result) => (
              <Stack key={result.id} tokens={{ childrenGap: 2 }} styles={{ root: { padding: '8px', backgroundColor: '#faf9f8', borderRadius: '4px' } }}>
                <div style={{ fontWeight: 600 }}>{result.data.Title || '(No Title)'}</div>
                <Label>{result.data.FileType} | {result.data.Author}</Label>
                <a href={result.data.Path} target="_blank" rel="noopener noreferrer" style={{ fontSize: '11px' }}>
                  {result.data.Path}
                </a>
              </Stack>
            ))
          ) : (
            <Label>Click &quot;Search All Content&quot; to see results.</Label>
          )}
        </Stack>
      </Stack>

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Faceted search with FileType and Author refiners. Click to filter/toggle.
      </Label>
    </Stack>
  );
};

/**
 * Example 6: Search Suggestions (Autocomplete)
 */
const PnPSearchSuggestionsExample: React.FC = () => {
  const [query, setQuery] = React.useState('');
  const [suggestions, setSuggestions] = React.useState<string[]>([]);
  const { suggest } = useSPFxPnPSearch();
  
  // Track if component is mounted to prevent state updates after unmount
  const mountedRef = React.useRef(true);
  
  // Cleanup on unmount
  React.useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
    };
  }, []);

  const handleSuggest = React.useCallback(async (value: string) => {
    if (!value || value.length < 2) {
      if (mountedRef.current) {
        setSuggestions([]);
      }
      return;
    }
    try {
      const results = await suggest(value);
      if (mountedRef.current) {
        setSuggestions(results);
      }
    } catch (err) {
      console.error('[PnPSearchSuggestions] Error:', err);
      if (mountedRef.current) {
        setSuggestions([]);
      }
    }
  }, [suggest]);

  // Debounced suggest with proper cleanup
  const debouncedSuggest = React.useMemo(() => {
    let timeoutId: number | undefined;
    
    // Return debounced function
    const debounced = (value: string): void => {
      if (timeoutId !== undefined) {
        clearTimeout(timeoutId);
      }
      timeoutId = window.setTimeout(() => handleSuggest(value), 300);
    };
    
    // Cleanup function stored in the debounced function
    debounced.cancel = (): void => {
      if (timeoutId !== undefined) {
        clearTimeout(timeoutId);
        timeoutId = undefined;
      }
    };
    
    return debounced;
  }, [handleSuggest]);
  
  // Cleanup timeout on unmount
  React.useEffect(() => {
    return () => {
      debouncedSuggest.cancel?.();
    };
  }, [debouncedSuggest]);

  const handleQueryChange = React.useCallback((value: string) => {
    setQuery(value);
    debouncedSuggest(value);
  }, [debouncedSuggest]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="SearchBookmark" style={{ marginRight: '8px' }} />
        Example 4: Search Suggestions (Autocomplete)
      </h3>
      <Separator />

      <TextField
        label="Search with Autocomplete"
        value={query}
        onChange={(_, newValue) => handleQueryChange(newValue ?? '')}
        placeholder="Type at least 2 characters..."
      />

      {suggestions.length > 0 ? (
        <Stack tokens={{ childrenGap: 2 }} styles={{ root: { padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
          <Label>Suggestions ({suggestions.length}):</Label>
          {suggestions.map((s, idx) => (
            <div
              key={idx}
              style={{ padding: '6px', backgroundColor: '#fff', borderRadius: '2px', cursor: 'pointer' }}
              onClick={() => setQuery(s)}
            >
              {s}
            </div>
          ))}
        </Stack>
      ) : query.length >= 2 ? (
        <MessageBar messageBarType={MessageBarType.info}>
          No suggestions available. This is normal in dev/test environments where there isn&apos;t enough indexed content.
          Try in a production environment with more search activity.
        </MessageBar>
      ) : null}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Real-time autocomplete with 300ms debounce. Requires indexed content and search activity.
      </Label>
    </Stack>
  );
};

/**
 * Example 7: useSPFxPnPList - CRUD Operations
 */
const PnPListExample: React.FC = () => {
  const [listTitle, setListTitle] = React.useState('');
  const [newTitle, setNewTitle] = React.useState('');
  const [editingId, setEditingId] = React.useState<number | null>(null);
  const [editTitle, setEditTitle] = React.useState('');

  const {
    query,
    items,
    loading,
    error,
    isEmpty,
    hasMore,
    loadMore,
    clearError,
    create,
    update,
    remove,
  } = useSPFxPnPList<{ Id: number; Title: string }>(listTitle, { pageSize: 10 });

  const handleLoadList = React.useCallback(() => {
    query(q => q.select('Id', 'Title').orderBy('Id', false)).catch(err => console.error('Load error:', err));
  }, [query]);

  const handleCreate = React.useCallback(async () => {
    if (!newTitle || !listTitle) return;
    try {
      await create({ Title: newTitle });
      setNewTitle('');
    } catch (err) {
      console.error('Create error:', err);
    }
  }, [create, newTitle, listTitle]);

  const handleUpdate = React.useCallback(async (id: number) => {
    if (!editTitle) return;
    try {
      await update(id, { Title: editTitle });
      setEditingId(null);
      setEditTitle('');
    } catch (err) {
      console.error('Update error:', err);
    }
  }, [update, editTitle]);

  const handleDelete = React.useCallback(async (id: number) => {
    if (!confirm('Delete this item?')) return;
    try {
      await remove(id);
    } catch (err) {
      console.error('Delete error:', err);
    }
  }, [remove]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="BulletedList" style={{ marginRight: '8px' }} />
        Example 3: useSPFxPnPList - CRUD Operations
      </h3>
      <Separator />

      <Stack horizontal tokens={{ childrenGap: 6 }}>
        <TextField
          label="List Title"
          value={listTitle}
          onChange={(_, newValue) => setListTitle(newValue ?? '')}
          placeholder="e.g., Site Pages"
          styles={{ root: { flexGrow: 1 } }}
        />
        <PrimaryButton
          text="Load"
          onClick={handleLoadList}
          disabled={!listTitle || loading}
          styles={{ root: { marginTop: '28px' } }}
        />
      </Stack>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      {listTitle && (
        <Stack horizontal tokens={{ childrenGap: 6 }} styles={{ root: { padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
          <TextField
            value={newTitle}
            onChange={(_, newValue) => setNewTitle(newValue ?? '')}
            placeholder="New item title..."
            styles={{ root: { flexGrow: 1 } }}
          />
          <PrimaryButton text="Create" onClick={handleCreate} disabled={!newTitle || loading} />
        </Stack>
      )}

      {loading && <MessageBar>Loading...</MessageBar>}
      {isEmpty && !loading && <MessageBar>No items found.</MessageBar>}

      {items.length > 0 && (
        <Stack tokens={{ childrenGap: 3 }}>
          <Label>Items ({items.length}):</Label>
          {items.map(item => (
            <Stack key={item.Id} horizontal tokens={{ childrenGap: 6 }} styles={{ root: { padding: '6px', backgroundColor: '#faf9f8', borderRadius: '4px' } }}>
              {editingId === item.Id ? (
                <>
                  <TextField value={editTitle} onChange={(_, v) => setEditTitle(v ?? '')} styles={{ root: { flexGrow: 1 } }} />
                  <DefaultButton text="Save" onClick={() => handleUpdate(item.Id)} />
                  <DefaultButton text="Cancel" onClick={() => { setEditingId(null); setEditTitle(''); }} />
                </>
              ) : (
                <>
                  <div style={{ flexGrow: 1 }}>#{item.Id} - {item.Title}</div>
                  <DefaultButton text="Edit" onClick={() => { setEditingId(item.Id); setEditTitle(item.Title); }} />
                  <DefaultButton text="Delete" onClick={() => handleDelete(item.Id)} />
                </>
              )}
            </Stack>
          ))}
          {hasMore && <PrimaryButton text="Load More" onClick={loadMore} disabled={loading} />}
        </Stack>
      )}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Complete CRUD with auto-refetch. Try with &quot;Site Pages&quot; or &quot;Documents&quot;.
      </Label>
    </Stack>
  );
};

// Helper function to safely stringify objects with circular references
const safeStringify = (obj: unknown, indent: number = 2): string => {
  const seen = new WeakSet();
  return JSON.stringify(obj, (key, value) => {
    // Skip private properties and service scopes to avoid circular references
    if (typeof key === 'string' && (key.indexOf('_') === 0 || key === 'serviceScope' || key === 'service')) {
      return '[Circular/Private]';
    }
    if (typeof value === 'object' && value !== null) {
      if (seen.has(value)) {
        return '[Circular Reference]';
      }
      seen.add(value);
    }
    return value;
  }, indent);
};

const SpFxReactToolkitTest: React.FC = () => {
  // Core Properties & Display
  const { properties, setProperties } = useSPFxProperties<IWebPartProps>();
  const { isEdit } = useSPFxDisplayMode();
  const { id, kind } = useSPFxInstanceInfo();

  // Theme & Environment
  const theme = useSPFxThemeInfo();
  const fluent9ThemeInfo = useSPFxFluent9ThemeInfo();
  const { type: envType } = useSPFxEnvironmentInfo();
  const isDarkTheme = theme?.isInverted ?? false;

  // User & Site Info
  const { displayName } = useSPFxUserInfo();
  const { title: siteTitle, webUrl, siteClassification } = useSPFxSiteInfo();
  const localeInfo = useSPFxLocaleInfo();

  // Teams Context
  const { supported: hasTeamsContext, theme: teamsTheme } = useSPFxTeams();

  // Page Context
  const pageContext = useSPFxPageContext();
  const pageTypeInfo = useSPFxPageType();

  // List & Hub Info (can be undefined)
  const listInfo = useSPFxListInfo();
  const hubInfo = useSPFxHubSiteInfo();

  // Container & Performance
  const containerSize = useSPFxContainerSize();
  const containerInfo = useSPFxContainerInfo();
  const performance = useSPFxPerformance();

  // Permissions
  const { hasWebPermission } = useSPFxPermissions();
  const canManageWeb = hasWebPermission(SPPermission.manageWeb);
  const canManageLists = hasWebPermission(SPPermission.manageLists);

  // Storage
  const sessionStorage = useSPFxSessionStorage('demo-session-key', '');
  const localStorage = useSPFxLocalStorage('demo-local-key', '');

  // Advanced
  const correlationInfo = useSPFxCorrelationInfo();
  const logger = useSPFxLogger();
  const serviceScope = useSPFxServiceScope();

  // HTTP Clients
  const spHttpClient = useSPFxSPHttpClient();
  const msGraphClient = useSPFxMSGraphClient();
  const aadHttpClient = useSPFxAadHttpClient();

  // OneDrive AppData hook (using instance id as folder for isolation)
  const oneDriveData = useSPFxOneDriveAppData<IOneDriveTestData>(
    'test-data.json',
    id, // Use instance ID as folder for multi-instance support
    false // Manual load for demo purposes
  );

  // Tenant Property hook (tenant-wide configuration)
  const tenantVersion = useSPFxTenantProperty<string>('spfx-toolkit-test-version', false);
  const tenantCounter = useSPFxTenantProperty<number>('spfx-toolkit-test-counter', false);

  // User Photo hook (current user profile photo)
  const userPhoto = useSPFxUserPhoto();

  // Local state for interactive demos
  const [descriptionInput, setDescriptionInput] = React.useState(properties?.description ?? '');
  const [sessionStorageInput, setSessionStorageInput] = React.useState('');
  const [localStorageInput, setLocalStorageInput] = React.useState('');
  const [oneDriveMessage, setOneDriveMessage] = React.useState('');
  const [tenantVersionInput, setTenantVersionInput] = React.useState('');
  const [showMessage, setShowMessage] = React.useState(false);
  const [messageText, setMessageText] = React.useState('');
  const [performanceResult, setPerformanceResult] = React.useState<string>('');
  const [logMessages, setLogMessages] = React.useState<Array<{ level: string; message: string }>>([]);
  const [crossSiteUrl, setCrossSiteUrl] = React.useState<string | undefined>(undefined);

  // Cross-site permissions (fetch only when URL is set)
  const crossSitePermissions = useSPFxCrossSitePermissions(crossSiteUrl);

  // Sync descriptionInput when properties change
  React.useEffect(() => {
    setDescriptionInput(properties?.description ?? '');
  }, [properties?.description]);

  // Handlers
  const handleUpdateProperties = React.useCallback(() => {
    setProperties({ description: descriptionInput });
    setMessageText('Properties updated successfully!');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [descriptionInput, setProperties]);

  const handleSaveSessionStorage = React.useCallback(() => {
    sessionStorage.setValue(sessionStorageInput);
    setSessionStorageInput('');
    setMessageText('Value saved to session storage!');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [sessionStorageInput, sessionStorage]);

  const handleSaveLocalStorage = React.useCallback(() => {
    localStorage.setValue(localStorageInput);
    setLocalStorageInput('');
    setMessageText('Value saved to local storage!');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [localStorageInput, localStorage]);

  const handlePerformanceTest = React.useCallback(async () => {
    const result = await performance.time('demo-test', async () => {
      // Simulate async operation
      await new Promise(resolve => setTimeout(resolve, 500));
      return 'Test completed';
    });
    setPerformanceResult(`${result.result} in ${result.durationMs.toFixed(2)}ms`);
  }, [performance]);

  const handleLog = React.useCallback((level: 'info' | 'warning' | 'error', message: string) => {
    if (level === 'info') logger.info(message);
    else if (level === 'warning') logger.warn(message);
    else logger.error(message);

    setLogMessages(prev => [...prev.slice(-4), { level, message }]);
  }, [logger]);

  const handleLoadOneDrive = React.useCallback(async () => {
    await oneDriveData.load();
    setMessageText('OneDrive load completed. Check status below (missing file sets isNotFound=true).');
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, [oneDriveData]);

  const handleSaveOneDrive = React.useCallback(async () => {
    try {
      const newData: IOneDriveTestData = {
        message: oneDriveMessage,
        counter: (oneDriveData.data?.counter ?? 0) + 1,
        timestamp: Date.now(),
      };
      await oneDriveData.write(newData);
      setOneDriveMessage('');
      setMessageText('Data saved to OneDrive successfully!');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Save failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [oneDriveMessage, oneDriveData]);

  const handleLoadTenantProperty = React.useCallback(async () => {
    try {
      await Promise.all([
        tenantVersion.load(),
        tenantCounter.load()
      ]);
      setMessageText('Tenant properties loaded successfully!');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Load failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantVersion, tenantCounter]);

  const handleSaveTenantVersion = React.useCallback(async () => {
    if (!tenantVersion.canWrite) {
      setMessageText('Insufficient permissions to write tenant properties');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
      return;
    }

    try {
      await tenantVersion.write(
        tenantVersionInput,
        'Test version property from SPFx React Toolkit'
      );
      setTenantVersionInput('');
      setMessageText('Version saved to tenant properties!');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Save failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantVersionInput, tenantVersion]);

  const handleIncrementTenantCounter = React.useCallback(async () => {
    if (!tenantCounter.canWrite) {
      setMessageText('Insufficient permissions to write tenant properties');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
      return;
    }

    try {
      const newValue = (tenantCounter.data ?? 0) + 1;
      await tenantCounter.write(
        newValue,
        'Test counter property from SPFx React Toolkit'
      );
      setMessageText(`Counter incremented to ${newValue}!`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Increment failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantCounter]);

  const handleRemoveTenantProperty = React.useCallback(async (propertyName: 'version' | 'counter') => {
    const hook = propertyName === 'version' ? tenantVersion : tenantCounter;
    
    if (!hook.canWrite) {
      setMessageText('Insufficient permissions to remove tenant properties');
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
      return;
    }

    if (!confirm(`Are you sure you want to remove the ${propertyName} property?`)) {
      return;
    }

    try {
      await hook.remove();
      setMessageText(`${propertyName} property removed!`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 3000);
    } catch (error) {
      setMessageText(`Remove failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setShowMessage(true);
      setTimeout(() => setShowMessage(false), 5000);
    }
  }, [tenantVersion, tenantCounter]);

  return (
    <section className={`${styles.spFxReactToolkitTest} ${hasTeamsContext ? styles.teams : ''}`}>
      {/* Message Bar */}
      {showMessage && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setShowMessage(false)}>
          {messageText}
        </MessageBar>
      )}

      {/* Main Content with Pivot Tabs */}
      <Pivot aria-label="SPFx React Toolkit Demo Tabs" style={{ marginTop: '20px' }}>

        {/* TAB 1: Overview */}
        <PivotItem headerText="Overview" itemIcon="Home">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: '16px' } }}>

            {/* Core Information Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Info" style={{ marginRight: '8px' }} />
                Core Information
                <Separator />
              </h3>
              <InfoRow label="Instance ID" value={id} icon="Fingerprint" />
              <InfoRow label="Component Kind" value={kind} icon="CubeShape" />
              <InfoRow label="Display Mode" value={isEdit ? 'Edit' : 'Read'} icon="Edit" />
              <InfoRow label="Environment Type" value={envType} icon="Globe" />
              <InfoRow label="Page Type" value={pageTypeInfo.pageType} icon="Page" />
            </Stack>

            {/* Properties Editor Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Settings" style={{ marginRight: '8px' }} />
                Properties Editor
                <Separator />
              </h3>
              <TextField
                label="Description Property"
                value={descriptionInput}
                onChange={(_, newValue) => setDescriptionInput(newValue ?? '')}
                placeholder="Enter description..."
              />
              <PrimaryButton
                text="Update Properties"
                onClick={handleUpdateProperties}
                iconProps={{ iconName: 'Save' }}
              />
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Changes will update the WebPart properties and trigger Property Pane refresh
              </Label>
            </Stack>

            {/* Container Info */}
            {containerSize && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="Resize" style={{ marginRight: '8px' }} />
                  Container Information
                  <Separator />
                </h3>
                <InfoRow label="Width" value={`${containerSize.width}px`} icon="ArrowRight" />
                <InfoRow label="Height" value={`${containerSize.height}px`} icon="ArrowUp" />
                <InfoRow label="DOM Element" value={containerInfo.element ? 'Available' : 'N/A'} icon="DOM" />
                <InfoRow label="Size Tracking" value={containerInfo.size ? 'Active' : 'Inactive'} icon="RadioBullet" />
              </Stack>
            )}

          </Stack>
        </PivotItem>

        {/* TAB 2: Context */}
        <PivotItem headerText="Context" itemIcon="Contact">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: '16px' } }}>

            {/* User & Site Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="People" style={{ marginRight: '8px' }} />
                User & Site Information
                <Separator />
              </h3>

              <InfoRow label="User Name" value={displayName} icon="Contact" />
              <InfoRow label="Site Title" value={siteTitle} icon="CityNext" />
              <InfoRow label="Site URL" value={webUrl} icon="Link" />
              {siteClassification && (
                <InfoRow label="Classification" value={siteClassification} icon="Tag" />
              )}
            </Stack>

            {/* Locale & Regional Settings */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="LocaleLanguage" style={{ marginRight: '8px' }} />
                Locale & Regional Settings
                <Separator />
              </h3>
              
              <InfoRow label="Content Locale" value={localeInfo.locale} icon="Globe" />
              <InfoRow label="UI Locale" value={localeInfo.uiLocale} icon="LocaleLanguage" />
              <InfoRow label="Text Direction" value={localeInfo.isRtl ? 'Right-to-Left (RTL)' : 'Left-to-Right (LTR)'} icon="TextAlignLeft" />
              
              {localeInfo.timeZone && (
                <>
                  <InfoRow label="Time Zone" value={localeInfo.timeZone.description} icon="Clock" />
                  <InfoRow label="UTC Offset" value={`${localeInfo.timeZone.offset} minutes`} icon="DateTime" />
                  <Label>
                    <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                    Current time: {new Intl.DateTimeFormat(localeInfo.locale, {
                      weekday: 'long',
                      year: 'numeric',
                      month: 'long',
                      day: 'numeric',
                      hour: '2-digit',
                      minute: '2-digit'
                    }).format(new Date())}
                  </Label>
                </>
              )}
              
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Use with JavaScript Intl APIs for i18n (dates, numbers, currencies)
              </Label>
            </Stack>

            {/* Teams Context */}
            {hasTeamsContext && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="TeamsLogo" style={{ marginRight: '8px' }} />
                  Microsoft Teams Context
                  <Separator />
                </h3>
                <InfoRow label="Teams Supported" value="Yes" icon="CompletedSolid" />
                <InfoRow label="Teams Theme" value={teamsTheme ?? 'default'} icon="Color" />
              </Stack>
            )}

            {/* List Context */}
            {listInfo && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="BulletedList" style={{ marginRight: '8px' }} />
                  List Context
                  <Separator />
                </h3>
                <InfoRow label="List ID" value={listInfo.id} icon="Fingerprint" />
                <InfoRow label="List Title" value={listInfo.title} icon="FabricTextHighlight" />
                {listInfo.baseTemplate && (
                  <InfoRow label="Base Template" value={String(listInfo.baseTemplate)} icon="PageData" />
                )}
                {listInfo.isDocumentLibrary && (
                  <Label>
                    <Icon iconName="FabricFolder" style={{ marginRight: '4px', color: '#0078d4' }} />
                    Document Library
                  </Label>
                )}
              </Stack>
            )}

            {/* Hub Site */}
            {hubInfo?.isHubSite && (
              <Stack tokens={{ childrenGap: 2 }}>
                <h3>
                  <Icon iconName="NetworkTower" style={{ marginRight: '8px' }} />
                  Hub Site Association
                  <Separator />
                </h3>
                
                {hubInfo.error && (
                  <MessageBar messageBarType={MessageBarType.error}>
                    Failed to load hub site URL: {hubInfo.error.message}
                  </MessageBar>
                )}
                
                <InfoRow label="Hub Site ID" value={hubInfo.hubSiteId} icon="Fingerprint" />
                
                {hubInfo.isLoading ? (
                  <MessageBar messageBarType={MessageBarType.info}>
                    Loading hub site URL...
                  </MessageBar>
                ) : hubInfo.hubSiteUrl ? (
                  <InfoRow label="Hub Site URL" value={hubInfo.hubSiteUrl} icon="Link" />
                ) : null}
              </Stack>
            )}

          </Stack>
        </PivotItem>

        {/* TAB 3: Advanced */}
        <PivotItem headerText="Advanced" itemIcon="DeveloperTools">
          <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: '16px' } }}>

            {/* Permissions Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Permissions" style={{ marginRight: '8px' }} />
                Permissions (Current Site)
                <Separator />
              </h3>
              <StatusBadge label="Can Manage Web" available={canManageWeb} />
              <StatusBadge label="Can Manage Lists" available={canManageLists} />
            </Stack>

            {/* Cross-Site Permissions Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Globe" style={{ marginRight: '8px' }} />
                Cross-Site Permissions
                <Separator />
              </h3>
              
              <Label>
                <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                Check permissions on a different site (auto-fetch when URL is provided)
              </Label>
              
              <TextField
                label="Target Site URL"
                value={crossSiteUrl || ''}
                onChange={(_, value) => setCrossSiteUrl(value?.trim() || undefined)}
                placeholder="https://contoso.sharepoint.com/sites/targetsite"
                description="Enter a site URL - permissions will be fetched automatically"
              />
              
              {crossSitePermissions.isLoading && (
                <MessageBar messageBarType={MessageBarType.info}>
                  Loading permissions from {crossSiteUrl}...
                </MessageBar>
              )}
              
              {crossSitePermissions.error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Error: {crossSitePermissions.error.message}
                </MessageBar>
              )}
              
              {crossSiteUrl && !crossSitePermissions.isLoading && !crossSitePermissions.error && (
                <Stack tokens={{ childrenGap: 1 }}>
                  <InfoRow 
                    label="Target Site" 
                    value={crossSiteUrl} 
                    icon="Globe" 
                  />
                  <StatusBadge 
                    label="Can Manage Web (Cross-Site)" 
                    available={crossSitePermissions.hasWebPermission(SPPermission.manageWeb)} 
                  />
                  <StatusBadge 
                    label="Can Manage Lists (Cross-Site)" 
                    available={crossSitePermissions.hasWebPermission(SPPermission.manageLists)} 
                  />
                  <StatusBadge 
                    label="Can Add Items (Cross-Site)" 
                    available={crossSitePermissions.hasWebPermission(SPPermission.addListItems)} 
                  />
                </Stack>
              )}
            </Stack>

            {/* User Photo Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="ContactCard" style={{ marginRight: '8px' }} />
                User Photo Demo
                <Separator />
              </h3>
              
              {userPhoto.error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Load Error: {userPhoto.error.message}
                </MessageBar>
              )}
              
              {userPhoto.isLoading ? (
                <MessageBar messageBarType={MessageBarType.info}>Loading photo from Microsoft Graph...</MessageBar>
              ) : userPhoto.photoUrl ? (
                <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                  <img 
                    src={userPhoto.photoUrl} 
                    alt={displayName}
                    style={{ 
                      width: 120, 
                      height: 120, 
                      borderRadius: '50%',
                      objectFit: 'cover',
                      border: '3px solid #0078d4',
                      boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
                    }} 
                  />
                  <Stack tokens={{ childrenGap: 1 }}>
                    <InfoRow label="Display Name" value={displayName} icon="Contact" />
                    <InfoRow label="Photo Size" value="240x240 (default)" icon="ResizeMouseMedium" />
                    <InfoRow label="Photo Format" value="Blob URL" icon="FileImage" />
                    <InfoRow label="Is Ready" value={userPhoto.isReady ? 'Yes' : 'No'} icon="CheckMark" />
                  </Stack>
                </Stack>
              ) : (
                <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                  <div style={{ 
                    width: 120, 
                    height: 120, 
                    borderRadius: '50%', 
                    backgroundColor: '#0078d4',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    color: 'white',
                    fontSize: 48,
                    fontWeight: 'bold'
                  }}>
                    {displayName ? displayName.charAt(0).toUpperCase() : '?'}
                  </div>
                  <Label>
                    <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                    No profile photo available. Showing initials fallback.
                  </Label>
                </Stack>
              )}
              
              <PrimaryButton 
                text={userPhoto.isLoading ? 'Loading...' : 'Reload Photo'} 
                onClick={userPhoto.reload} 
                disabled={userPhoto.isLoading}
                iconProps={{ iconName: 'Refresh' }} 
              />
              
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Photo loaded from Microsoft Graph API (/me/photos/240x240/$value)
              </Label>
              <Label>
                <Icon iconName="Warning" style={{ marginRight: '4px', color: '#d83b01' }} />
                Requires Microsoft Graph permissions: User.Read
              </Label>
              <Label>
                <Icon iconName="LightningBolt" style={{ marginRight: '4px', color: '#107c10' }} />
                Blob URL automatically cleaned up on unmount (memory safe)
              </Label>
            </Stack>

            {/* OneDrive AppData Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="OneDrive" style={{ marginRight: '8px' }} />
                OneDrive AppData Demo
                <Separator />
              </h3>
              
              {oneDriveData.error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Load Error: {oneDriveData.error.message}
                </MessageBar>
              )}
              
              {oneDriveData.writeError && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Write Error: {oneDriveData.writeError.message}
                </MessageBar>
              )}

              {!oneDriveData.isLoading && oneDriveData.isNotFound && !oneDriveData.data && !oneDriveData.error && (
                <MessageBar messageBarType={MessageBarType.info}>
                  File not found in OneDrive yet (isNotFound=true). Click &quot;Save to OneDrive&quot; to create it.
                </MessageBar>
              )}
              
              {oneDriveData.isLoading ? (
                <MessageBar messageBarType={MessageBarType.info}>Loading from OneDrive...</MessageBar>
              ) : oneDriveData.data ? (
                <Stack tokens={{ childrenGap: 1 }}>
                  <InfoRow label="Message" value={oneDriveData.data.message} icon="Message" />
                  <InfoRow label="Counter" value={String(oneDriveData.data.counter)} icon="NumberField" />
                  <InfoRow label="Last Updated" value={new Date(oneDriveData.data.timestamp).toLocaleString()} icon="DateTime" />
                  <InfoRow label="Is Ready" value={oneDriveData.isReady ? 'Yes' : 'No'} icon="CheckMark" />
                  <InfoRow label="Is Not Found" value={oneDriveData.isNotFound ? 'Yes' : 'No'} icon="BlockedSiteSolid12" />
                </Stack>
              ) : (
                <Label>
                  {oneDriveData.isNotFound
                    ? 'No file found yet. Click "Save to OneDrive" to create it.'
                    : 'No data loaded yet. Click "Load" to fetch from OneDrive.'}
                </Label>
              )}
              
              <TextField
                label="New Message"
                value={oneDriveMessage}
                onChange={(_, newValue) => setOneDriveMessage(newValue ?? '')}
                placeholder="Enter a message to save..."
                disabled={oneDriveData.isWriting || oneDriveData.isLoading}
              />
              
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <PrimaryButton 
                  text="Load from OneDrive" 
                  onClick={handleLoadOneDrive} 
                  disabled={oneDriveData.isLoading || oneDriveData.isWriting}
                  iconProps={{ iconName: 'CloudDownload' }} 
                />
                <PrimaryButton 
                  text={oneDriveData.isWriting ? 'Saving...' : 'Save to OneDrive'} 
                  onClick={handleSaveOneDrive} 
                  disabled={!oneDriveMessage || oneDriveData.isWriting || oneDriveData.isLoading}
                  iconProps={{ iconName: 'CloudUpload' }} 
                />
              </Stack>
              
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Data is stored in your OneDrive (appRoot:/{id}/test-data.json). Each WebPart instance has its own file.
              </Label>
              <Label>
                <Icon iconName="Warning" style={{ marginRight: '4px', color: '#d83b01' }} />
                Requires Microsoft Graph permissions: Files.ReadWrite or Files.ReadWrite.AppFolder
              </Label>
            </Stack>

            {/* Tenant Property Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Org" style={{ marginRight: '8px' }} />
                Tenant Property Demo
                <Separator />
              </h3>
              
              {(tenantVersion.error || tenantCounter.error) && (
                <MessageBar messageBarType={MessageBarType.error}>
                  Load Error: {tenantVersion.error?.message || tenantCounter.error?.message}
                </MessageBar>
              )}
              
              {(tenantVersion.writeError || tenantCounter.writeError) && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Write Error: {tenantVersion.writeError?.message || tenantCounter.writeError?.message}
                </MessageBar>
              )}
              
              {(tenantVersion.isLoading || tenantCounter.isLoading) ? (
                <MessageBar messageBarType={MessageBarType.info}>Loading from tenant app catalog...</MessageBar>
              ) : (tenantVersion.data || tenantCounter.data) ? (
                <Stack tokens={{ childrenGap: 1 }}>
                  <InfoRow label="Version" value={tenantVersion.data ?? '(not set)'} icon="ServerEnviroment" />
                  {tenantVersion.description && (
                    <Label>
                      <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                      {tenantVersion.description}
                    </Label>
                  )}
                  <InfoRow label="Counter" value={String(tenantCounter.data ?? 0)} icon="NumberField" />
                  {tenantCounter.description && (
                    <Label>
                      <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                      {tenantCounter.description}
                    </Label>
                  )}
                  <InfoRow label="Can Write" value={tenantVersion.canWrite ? 'Yes' : 'No'} icon="Permissions" />
                </Stack>
              ) : (
                <Label>No data loaded yet. Click &quot;Load&quot; to fetch from tenant properties.</Label>
              )}
              
              <TextField
                label="New Version"
                value={tenantVersionInput}
                onChange={(_, newValue) => setTenantVersionInput(newValue ?? '')}
                placeholder="e.g., 1.0.0"
                disabled={tenantVersion.isWriting || tenantVersion.isLoading || !tenantVersion.canWrite}
              />
              
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <PrimaryButton 
                  text="Load Properties" 
                  onClick={handleLoadTenantProperty} 
                  disabled={tenantVersion.isLoading || tenantCounter.isLoading}
                  iconProps={{ iconName: 'CloudDownload' }} 
                />
                <PrimaryButton 
                  text={tenantVersion.isWriting ? 'Saving...' : 'Save Version'} 
                  onClick={handleSaveTenantVersion} 
                  disabled={!tenantVersionInput || tenantVersion.isWriting || !tenantVersion.canWrite}
                  iconProps={{ iconName: 'Save' }} 
                />
                <PrimaryButton 
                  text={tenantCounter.isWriting ? 'Incrementing...' : 'Increment Counter'} 
                  onClick={handleIncrementTenantCounter} 
                  disabled={tenantCounter.isWriting || !tenantCounter.canWrite}
                  iconProps={{ iconName: 'Add' }} 
                />
              </Stack>
              
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <DefaultButton 
                  text="Remove Version" 
                  onClick={() => handleRemoveTenantProperty('version')} 
                  disabled={tenantVersion.isWriting || !tenantVersion.canWrite || !tenantVersion.data}
                  iconProps={{ iconName: 'Delete' }} 
                />
                <DefaultButton 
                  text="Remove Counter" 
                  onClick={() => handleRemoveTenantProperty('counter')} 
                  disabled={tenantCounter.isWriting || !tenantCounter.canWrite || !tenantCounter.data}
                  iconProps={{ iconName: 'Delete' }} 
                />
              </Stack>
              
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Properties are stored tenant-wide in SharePoint StorageEntity. All users can read, only admins can write.
              </Label>
              <Label>
                <Icon iconName="Info" style={{ marginRight: '4px', color: '#0078d4' }} />
                Keys: spfx-toolkit-test-version (string), spfx-toolkit-test-counter (number)
              </Label>
              <Label>
                <Icon iconName="Warning" style={{ marginRight: '4px', color: '#d83b01' }} />
                Requires: Tenant app catalog provisioned, Manage Web permissions to write
              </Label>
              
              {!tenantVersion.canWrite && (
                <MessageBar messageBarType={MessageBarType.info}>
                  ℹ️ You don&apos;t have permission to modify tenant properties. Contact your SharePoint administrator.
                </MessageBar>
              )}
            </Stack>


            {/* Storage Demo Cards */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Database" style={{ marginRight: '8px' }} />
                Session Storage Demo
                <Separator />
              </h3>

              <InfoRow label="Current Value" value={sessionStorage.value || '(empty)'} icon="Variable" />
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <TextField
                  value={sessionStorageInput}
                  onChange={(_, newValue) => setSessionStorageInput(newValue ?? '')}
                  placeholder="Enter value..."
                  styles={{ root: { flexGrow: 1 } }}
                />
                <PrimaryButton text="Save" onClick={handleSaveSessionStorage} iconProps={{ iconName: 'Save' }} />
                <DefaultButton text="Clear" onClick={() => sessionStorage.remove()} iconProps={{ iconName: 'Delete' }} />
              </Stack>
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Persists only for current tab/session
              </Label>
            </Stack>

            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Save" style={{ marginRight: '8px' }} />
                Local Storage Demo
                <Separator />
              </h3>
              <InfoRow label="Current Value" value={localStorage.value || '(empty)'} icon="Variable" />
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <TextField
                  value={localStorageInput}
                  onChange={(_, newValue) => setLocalStorageInput(newValue ?? '')}
                  placeholder="Enter value..."
                  styles={{ root: { flexGrow: 1 } }}
                />
                <PrimaryButton text="Save" onClick={handleSaveLocalStorage} iconProps={{ iconName: 'Save' }} />
                <DefaultButton text="Clear" onClick={() => localStorage.remove()} iconProps={{ iconName: 'Delete' }} />
              </Stack>
              <Label>
                <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
                Persists across sessions and page reloads
              </Label>
            </Stack>

            {/* Performance Test Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="SpeedHigh" style={{ marginRight: '8px' }} />
                Performance Test
                <Separator />
              </h3>
              {performanceResult && (
                <MessageBar messageBarType={MessageBarType.info}>
                  {performanceResult}
                </MessageBar>
              )}
              <PrimaryButton
                text="Run Performance Test"
                onClick={handlePerformanceTest}
                iconProps={{ iconName: 'LightningBolt' }}
              />
            </Stack>

            {/* Logger Demo Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Code" style={{ marginRight: '8px' }} />
                Logger Demo
                <Separator />
              </h3>
              <Stack horizontal tokens={{ childrenGap: 2 }}>
                <DefaultButton
                  text="Info"
                  onClick={() => handleLog('info', 'This is an info message')}
                  iconProps={{ iconName: 'Info' }}
                />
                <DefaultButton
                  text="Warning"
                  onClick={() => handleLog('warning', 'This is a warning message')}
                  iconProps={{ iconName: 'Warning' }}
                />
                <DefaultButton
                  text="Error"
                  onClick={() => handleLog('error', 'This is an error message')}
                  iconProps={{ iconName: 'Error' }}
                />
              </Stack>
              {logMessages.length > 0 && (
                <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginTop: '8px' } }}>
                  <Label>Recent Logs (check browser console):</Label>
                  {logMessages.map((log, idx) => {
                    const levelClass = log.level === 'info' ? styles.info :
                      log.level === 'warning' ? styles.warning :
                        log.level === 'error' ? styles.error :
                          styles.verbose;
                    return (
                      <div key={idx} className={`${styles.logMessage} ${levelClass}`}>
                        <strong>[{log.level.toUpperCase()}]</strong>: {log.message}
                      </div>
                    );
                  })}
                </Stack>
              )}
            </Stack>

            {/* HTTP Clients Status Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="CloudUpload" style={{ marginRight: '8px' }} />
                HTTP Clients Status
                <Separator />
              </h3>
              <StatusBadge label="SPHttpClient" available={!!spHttpClient} />
              <StatusBadge label="MSGraphClient" available={!!msGraphClient} />
              <StatusBadge label="AadHttpClient" available={!!aadHttpClient} />
            </Stack>

            {/* Theme Information Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Color" style={{ marginRight: '8px' }} />
                Theme Information
                <Separator />
              </h3>
              <InfoRow label="Is Dark Theme" value={isDarkTheme ? 'Yes' : 'No'} icon="Brightness" />
              <InfoRow label="Body Background" value={theme?.semanticColors?.bodyBackground ?? 'N/A'} icon="FabricFolderFill" />
              <InfoRow label="Body Text" value={theme?.semanticColors?.bodyText ?? 'N/A'} icon="Font" />
              <InfoRow label="Link Color" value={theme?.semanticColors?.link ?? 'N/A'} icon="Link" />
            </Stack>

            {/* Fluent UI 9 Theme Information Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="Color" style={{ marginRight: '8px' }} />
                Fluent UI 9 Theme Information
                <Separator />
              </h3>
              <InfoRow label="Is Teams Context" value={fluent9ThemeInfo.isTeams ? 'Yes' : 'No'} icon="TeamsLogo" />
              {fluent9ThemeInfo.isTeams && fluent9ThemeInfo.teamsTheme && (
                <InfoRow label="Teams Theme" value={fluent9ThemeInfo.teamsTheme} icon="Color" />
              )}
              <details className={styles.detailsSection}>
                <summary>
                  Click to expand Fluent UI 9 Theme (JSON)
                </summary>
                <pre>
                  {safeStringify(fluent9ThemeInfo.theme, 2)}
                </pre>
              </details>
            </Stack>

            {/* Advanced Diagnostics Card */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="DeveloperTools" style={{ marginRight: '8px' }} />
                Advanced Diagnostics
                <Separator />
              </h3>
              <InfoRow label="Correlation ID" value={correlationInfo.correlationId} icon="TrackersMirrored" />
              <InfoRow label="Tenant ID" value={correlationInfo.tenantId} icon="CityNext" />
              <InfoRow label="ServiceScope" value={serviceScope ? 'Available' : 'N/A'} icon="Settings" />
            </Stack>

            {/* Page Context Raw Data */}
            <Stack tokens={{ childrenGap: 2 }}>
              <h3>
                <Icon iconName="FileCode" style={{ marginRight: '8px' }} />
                Page Context (Raw JSON)
                <Separator />
              </h3>
              <details className={styles.detailsSection}>
                <summary>
                  Click to expand/collapse
                </summary>
                <pre>
                  {safeStringify(pageContext, 2)}
                </pre>
              </details>
            </Stack>

          </Stack>
        </PivotItem>

        {/* TAB 4: HttpClient Example */}
        <PivotItem headerText="HttpClient" itemIcon="Cloud">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: '16px' } }}>

            <MessageBar messageBarType={MessageBarType.info}>
              <strong>useSPFxHttpClient Hook Example</strong>
              <br />
              This hook provides access to the generic HttpClient for calling external APIs (non-SharePoint).
              For SharePoint REST API calls, use <strong>useSPFxSPHttpClient</strong> instead.
            </MessageBar>

            <HttpClientExample />

          </Stack>
        </PivotItem>

        {/* TAB 5: PnPjs Examples */}
        <PivotItem headerText="PnPjs" itemIcon="CloudDownload">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: '16px' } }}>

            {/* Example 1: useSPFxPnPContext */}
            <PnPContextExample />

            {/* Example 2: useSPFxPnP */}
            <PnPOperationsExample />

            {/* Example 3: useSPFxPnPList */}
            <PnPListExample />

          </Stack>
        </PivotItem>

        {/* TAB 6: Search Examples */}
        <PivotItem headerText="Search" itemIcon="Search">
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: '16px' } }}>

            <MessageBar messageBarType={MessageBarType.info}>
              <strong>useSPFxPnPSearch Hook Examples</strong>
              <br />
              These examples demonstrate SharePoint Search capabilities using the native PnPjs SearchQueryBuilder API.
            </MessageBar>

            {/* Example 1: Basic Search */}
            <PnPSearchBasicExample />

            {/* Example 2: Advanced Search with Builder */}
            <PnPSearchAdvancedExample />

            {/* Example 3: Faceted Search with Refiners */}
            <PnPSearchRefinersExample />

            {/* Example 4: Search Suggestions (Autocomplete) */}
            <PnPSearchSuggestionsExample />

          </Stack>
        </PivotItem>

      </Pivot>

    </section>
  );
};

export default SpFxReactToolkitTest;
