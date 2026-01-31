# PnPjs Hooks

> Hooks for PnPjs v4 integration with state management

## Overview

These hooks provide PnPjs v4 integration with automatic state management, batching support, and type-safe query building.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxPnP`](#usespfxpnp) | `SPFxPnPInfo` | PnPjs with invoke/batch helpers |
| [`useSPFxPnPContext`](#usespfxpnpcontext) | `PnPContextInfo` | PnPjs SPFI factory |
| [`useSPFxPnPList`](#usespfxpnplist) | `SPFxPnPListInfo<T>` | List CRUD operations |
| [`useSPFxPnPSearch`](#usespfxpnpsearch) | `SPFxPnPSearchInfo<T>` | SharePoint Search |

### Selective Imports

PnPjs uses selective imports for tree-shaking. Import only what you need:

```typescript
// For lists and items
import '@pnp/sp/lists';
import '@pnp/sp/items';

// For files and folders
import '@pnp/sp/files';
import '@pnp/sp/folders';

// For search
import '@pnp/sp/search';

// For user profiles
import '@pnp/sp/profiles';
```

---

## useSPFxPnP

Access PnPjs with state management and batching support.

### Signature

```typescript
function useSPFxPnP(pnpContext?: PnPContextInfo): SPFxPnPInfo
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `pnpContext` | `PnPContextInfo` | No | Optional PnPjs context (default: auto-created) |

### Returns

```typescript
interface SPFxPnPInfo {
  /** Configured SPFI instance for direct access */
  readonly sp: SPFI | undefined;
  
  /** Execute single operation with state management */
  readonly invoke: <T>(fn: (sp: SPFI) => Promise<T>) => Promise<T>;
  
  /** Execute multiple operations in one batch request */
  readonly batch: <T>(fn: (batchedSP: SPFI) => Promise<T>) => Promise<T>;
  
  /** Loading state during invoke/batch calls */
  readonly isLoading: boolean;
  
  /** Last error from operations */
  readonly error: Error | undefined;
  
  /** Clear error state */
  readonly clearError: () => void;
  
  /** True if sp instance is initialized */
  readonly isInitialized: boolean;
  
  /** Effective site URL being used */
  readonly siteUrl: string;
}
```

### Description

Provides convenient wrappers around PnPjs SPFI instance with automatic loading and error state tracking.

**Three usage patterns:**
1. **invoke()** - Single operations with automatic state management
2. **batch()** - Multiple operations in one request with state management
3. **sp** - Direct access for full control (advanced scenarios)

### Example: Basic Query

```tsx
import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
import '@pnp/sp/lists';
import '@pnp/sp/items';

function TaskList() {
  const { invoke, isLoading, error } = useSPFxPnP();
  const [tasks, setTasks] = React.useState([]);
  
  const loadTasks = () => {
    invoke(sp => 
      sp.web.lists
        .getByTitle('Tasks')
        .items
        .select('Id', 'Title', 'Status')
        .filter("Status eq 'Active'")
        .orderBy('Created', false)
        .top(50)()
    ).then(setTasks);
  };
  
  React.useEffect(() => { loadTasks(); }, []);
  
  if (isLoading) return <Spinner />;
  if (error) return <ErrorMessage message={error.message} />;
  
  return (
    <ul>
      {tasks.map(task => <li key={task.Id}>{task.Title}</li>)}
    </ul>
  );
}
```

### Example: Batch Operations

```tsx
import { useSPFxPnP } from '@apvee/spfx-react-toolkit';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/webs';

function Dashboard() {
  const { batch, isLoading } = useSPFxPnP();
  const [data, setData] = React.useState({ lists: [], user: null, tasks: [] });
  
  const loadDashboard = async () => {
    // All requests executed in single HTTP call
    const [lists, user, tasks] = await batch(async (batchedSP) => {
      const lists = batchedSP.web.lists();
      const user = batchedSP.web.currentUser();
      const tasks = batchedSP.web.lists.getByTitle('Tasks').items.top(10)();
      
      return Promise.all([lists, user, tasks]);
    });
    
    setData({ lists, user, tasks });
  };
  
  React.useEffect(() => { loadDashboard(); }, []);
  
  return (
    <div>
      <h2>Welcome, {data.user?.Title}</h2>
      <p>You have access to {data.lists.length} lists</p>
      <h3>Recent Tasks</h3>
      <ul>
        {data.tasks.map(t => <li key={t.Id}>{t.Title}</li>)}
      </ul>
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxPnP.ts)

---

## useSPFxPnPContext

Access PnPjs SPFI factory for custom configuration.

### Signature

```typescript
function useSPFxPnPContext(config?: PnPContextConfig): PnPContextInfo
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `config` | `PnPContextConfig` | No | Custom PnPjs configuration |

### Config Options

```typescript
interface PnPContextConfig {
  /** Target site URL (default: current site) */
  siteUrl?: string;
  
  /** Enable caching (default: false) */
  enableCaching?: boolean;
  
  /** Cache timeout in ms (default: 60000) */
  cacheTimeout?: number;
}
```

### Returns

```typescript
interface PnPContextInfo {
  /** Configured SPFI instance */
  readonly sp: SPFI | undefined;
  
  /** Current site URL */
  readonly siteUrl: string;
  
  /** Initialization error if any */
  readonly error: Error | undefined;
  
  /** Whether initialization is complete */
  readonly isInitialized: boolean;
}
```

### Example: Custom Site Context

```tsx
import { useSPFxPnPContext, useSPFxPnP } from '@apvee/spfx-react-toolkit';

function CrossSiteData() {
  // Create context for another site
  const otherSiteContext = useSPFxPnPContext({
    siteUrl: 'https://tenant.sharepoint.com/sites/OtherSite'
  });
  
  // Use that context with useSPFxPnP
  const { invoke, isLoading } = useSPFxPnP(otherSiteContext);
  const [items, setItems] = React.useState([]);
  
  const loadItems = () => {
    invoke(sp => sp.web.lists.getByTitle('Documents').items())
      .then(setItems);
  };
  
  return (
    <button onClick={loadItems} disabled={isLoading}>
      Load from Other Site
    </button>
  );
}
```

### Example: With Caching

```tsx
import { useSPFxPnPContext, useSPFxPnP } from '@apvee/spfx-react-toolkit';

function CachedData() {
  const context = useSPFxPnPContext({
    enableCaching: true,
    cacheTimeout: 300000 // 5 minutes
  });
  
  const { invoke } = useSPFxPnP(context);
  
  // Repeated calls within 5 minutes will use cached data
  const loadData = () => invoke(sp => sp.web.lists());
}
```

### Source

[View source](../../src/hooks/useSPFxPnPContext.ts)

---

## useSPFxPnPList

Full CRUD operations for SharePoint lists with pagination.

### Signature

```typescript
function useSPFxPnPList<T = unknown>(
  listTitle: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T>
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `listTitle` | `string` | Yes | Title of the SharePoint list |
| `options` | `UseSPFxPnPListOptions` | No | Configuration options |
| `pnpContext` | `PnPContextInfo` | No | Optional custom PnP context |

### Options

```typescript
interface UseSPFxPnPListOptions {
  /** Page size for pagination (default: 100) */
  pageSize?: number;
}
```

### Returns

```typescript
interface SPFxPnPListInfo<T = unknown> {
  /** Execute query with fluent API */
  query: (queryBuilder?: (items: IItems) => IItems) => Promise<T[]>;
  
  /** Current items from last query */
  items: T[];
  
  /** Loading state */
  loading: boolean;
  
  /** Loading more items state */
  loadingMore: boolean;
  
  /** Current error */
  error: Error | undefined;
  
  /** Whether items array is empty */
  isEmpty: boolean;
  
  /** Whether more items available */
  hasMore: boolean;
  
  /** Re-execute last query */
  refetch: () => Promise<void>;
  
  /** Load next page of items */
  loadMore: () => Promise<T[]>;
  
  /** Clear error state */
  clearError: () => void;
  
  /** Get single item by ID */
  getById: (id: number) => Promise<T | undefined>;
  
  /** Create new item */
  create: (item: Partial<T>) => Promise<number>;
  
  /** Update item by ID */
  update: (id: number, item: Partial<T>) => Promise<void>;
  
  /** Delete item by ID */
  remove: (id: number) => Promise<void>;
  
  /** Batch create multiple items */
  batchCreate: (items: Partial<T>[]) => Promise<number[]>;
  
  /** Batch update multiple items */
  batchUpdate: (updates: Array<{ id: number; item: Partial<T> }>) => Promise<void>;
  
  /** Batch delete multiple items */
  batchDelete: (ids: number[]) => Promise<void>;
}
```

### Example: Complete CRUD

```tsx
import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';
import '@pnp/sp/lists';
import '@pnp/sp/items';

interface ITask {
  Id: number;
  Title: string;
  Status: string;
  Priority: number;
}

function TaskManager() {
  const { 
    query, 
    items, 
    loading, 
    hasMore,
    loadMore,
    create, 
    update, 
    remove,
    refetch,
    error 
  } = useSPFxPnPList<ITask>('Tasks', { pageSize: 25 });
  
  // Load initial data
  React.useEffect(() => {
    query(q => 
      q.select('Id', 'Title', 'Status', 'Priority')
       .filter("Status ne 'Completed'")
       .orderBy('Priority', false)
    );
  }, []);
  
  // Create new task
  const addTask = async (title: string) => {
    await create({ Title: title, Status: 'New', Priority: 1 });
    // refetch() is called automatically
  };
  
  // Update task status
  const completeTask = async (id: number) => {
    await update(id, { Status: 'Completed' });
  };
  
  // Delete task
  const deleteTask = async (id: number) => {
    await remove(id);
  };
  
  if (loading) return <Spinner />;
  if (error) return <ErrorMessage message={error.message} />;
  
  return (
    <div>
      <AddTaskForm onAdd={addTask} />
      <ul>
        {items.map(task => (
          <li key={task.Id}>
            {task.Title} - {task.Status}
            <button onClick={() => completeTask(task.Id)}>✓</button>
            <button onClick={() => deleteTask(task.Id)}>×</button>
          </li>
        ))}
      </ul>
      {hasMore && (
        <button onClick={loadMore}>Load More</button>
      )}
    </div>
  );
}
```

### Example: Batch Operations

```tsx
import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';

function BulkOperations() {
  const { batchCreate, batchUpdate, batchDelete } = useSPFxPnPList('Tasks');
  
  // Create multiple items in one request
  const createBulk = async () => {
    const newIds = await batchCreate([
      { Title: 'Task 1', Status: 'New' },
      { Title: 'Task 2', Status: 'New' },
      { Title: 'Task 3', Status: 'New' }
    ]);
    console.log('Created IDs:', newIds);
  };
  
  // Update multiple items in one request
  const updateBulk = async (ids: number[]) => {
    await batchUpdate(
      ids.map(id => ({ id, item: { Status: 'Completed' } }))
    );
  };
  
  // Delete multiple items in one request
  const deleteBulk = async (ids: number[]) => {
    await batchDelete(ids);
  };
}
```

### Source

[View source](../../src/hooks/useSPFxPnPList.ts)

---

## useSPFxPnPSearch

SharePoint Search with pagination, refiners, and suggestions.

### Signature

```typescript
function useSPFxPnPSearch<T = Record<string, string>>(
  options?: UseSPFxPnPSearchOptions
): SPFxPnPSearchInfo<T>
```

### Options

```typescript
interface UseSPFxPnPSearchOptions {
  /** Results per page (default: 50) */
  pageSize?: number;
  
  /** Default properties to select */
  selectProperties?: string[];
  
  /** Refiners to request (comma-separated) */
  refiners?: string;
}
```

### Returns

```typescript
interface SPFxPnPSearchInfo<T> {
  /** Execute search query */
  search: (queryBuilder: SearchQueryBuilderFn) => Promise<SearchResult<T>[]>;
  
  /** Get search suggestions */
  suggest: (query: string) => Promise<string[]>;
  
  /** Current search results */
  results: SearchResult<T>[];
  
  /** Current refiners/facets */
  refiners: SearchRefiner[];
  
  /** Total results count */
  totalRows: number;
  
  /** Loading state */
  loading: boolean;
  
  /** Load more results */
  loadMore: () => Promise<SearchResult<T>[]>;
  
  /** Has more results */
  hasMore: boolean;
  
  /** Current error */
  error: Error | undefined;
  
  /** Clear error */
  clearError: () => void;
}
```

### Search Verticals

Pre-defined result sources for filtering:

```typescript
import { SearchVerticals } from '@apvee/spfx-react-toolkit';

SearchVerticals.All         // All results (default)
SearchVerticals.People      // People and profiles
SearchVerticals.Videos      // Video content
SearchVerticals.Sites       // SharePoint sites
SearchVerticals.Documents   // Documents
SearchVerticals.Conversations // Teams/Yammer
SearchVerticals.Pages       // SharePoint pages
```

### Example: Basic Search

```tsx
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
import '@pnp/sp/search';

function SearchBox() {
  const { search, results, loading, totalRows, hasMore, loadMore } = useSPFxPnPSearch({
    pageSize: 20,
    selectProperties: ['Title', 'Path', 'Author', 'LastModifiedTime']
  });
  
  const [query, setQuery] = React.useState('');
  
  const doSearch = () => {
    search(builder => 
      builder.text(query)
    );
  };
  
  return (
    <div>
      <input 
        value={query} 
        onChange={e => setQuery(e.target.value)} 
        placeholder="Search..."
      />
      <button onClick={doSearch} disabled={loading}>Search</button>
      
      <p>Found {totalRows} results</p>
      
      <ul>
        {results.map(result => (
          <li key={result.id}>
            <a href={result.data.Path}>{result.data.Title}</a>
            <span>by {result.data.Author}</span>
          </li>
        ))}
      </ul>
      
      {hasMore && (
        <button onClick={loadMore}>Load More</button>
      )}
    </div>
  );
}
```

### Example: With Refiners

```tsx
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';

function SearchWithFilters() {
  const { search, results, refiners, loading } = useSPFxPnPSearch({
    refiners: 'FileType,Author,ModifiedBy'
  });
  
  const [selectedFileType, setSelectedFileType] = React.useState<string | null>(null);
  
  const doSearch = (query: string) => {
    search(builder => {
      let b = builder.text(query);
      
      if (selectedFileType) {
        b = b.refinementFilters(`FileType:equals("${selectedFileType}")`);
      }
      
      return b;
    });
  };
  
  return (
    <div>
      <div className="filters">
        <h3>File Types</h3>
        {refiners.find(r => r.name === 'FileType')?.entries.map(entry => (
          <button 
            key={entry.value}
            onClick={() => setSelectedFileType(entry.value)}
          >
            {entry.value} ({entry.count})
          </button>
        ))}
      </div>
      
      <div className="results">
        {results.map(r => <ResultCard key={r.id} result={r} />)}
      </div>
    </div>
  );
}
```

### Example: People Search

```tsx
import { useSPFxPnPSearch, SearchVerticals } from '@apvee/spfx-react-toolkit';

function PeopleSearch() {
  const { search, results, loading } = useSPFxPnPSearch({
    selectProperties: ['AccountName', 'PreferredName', 'Department', 'JobTitle', 'PictureURL']
  });
  
  const searchPeople = (query: string) => {
    search(builder => 
      builder.text(query)
        .sourceId(SearchVerticals.People)
    );
  };
  
  return (
    <div>
      {results.map(person => (
        <Persona
          key={person.id}
          imageUrl={person.data.PictureURL}
          text={person.data.PreferredName}
          secondaryText={person.data.JobTitle}
          tertiaryText={person.data.Department}
        />
      ))}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxPnPSearch.ts)

---

## See Also

- [HTTP Client Hooks](./http-clients.md) - Native SPFx HTTP clients
- [Context Hooks](./context.md) - Context access
- [PnPjs Documentation](https://pnp.github.io/pnpjs/) - Official PnPjs docs

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
