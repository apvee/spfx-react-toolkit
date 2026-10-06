import type { PnPContextInfo } from './useSPFxPnPContext';
import { useSPFxPnPListInternal } from './useSPFxPnPList.internal';

// Import PnPjs native types for proper type safety
import type { IItems } from '@pnp/sp/items';
import type { InitialFieldQuery, ComparisonResult } from '@pnp/sp/spqueryable';

/**
 * Type representing a fluent filter function for type-safe query building.
 * Uses native PnPjs v4 fluent filter API types.
 * 
 * @template T - The type of the list item
 * @example
 * ```tsx
 * interface Task {
 *   Status: string;
 *   Priority: number;
 * }
 * 
 * const filterFn: ListFilterFunction<Task> = (f) => 
 *   f.text("Status").equals("Active")
 *     .and()
 *     .number("Priority").greaterThan(3);
 * ```
 */
export type ListFilterFunction<T> = (f: InitialFieldQuery<T>) => ComparisonResult<T>;

/**
 * Options for configuring the useSPFxPnPList hook.
 */
export interface UseSPFxPnPListOptions {
  /**
   * Page size for pagination (number of items per page).
   * Used by `loadMore()` for automatic pagination.
   * 
   * @default 100
   * @example
   * ```tsx
   * const { query, loadMore } = useSPFxPnPList('Tasks', { pageSize: 50 });
   * ```
   */
  pageSize?: number;
}

/**
 * Return type for the useSPFxPnPList hook.
 * 
 * @template T - The type of the list item
 */
export interface SPFxPnPListInfo<T = unknown> {
  /**
   * Executes a query against the SharePoint list using PnPjs fluent API.
   * 
   * The hook automatically detects if `.top()` is called in the queryBuilder:
   * - If `.top()` is specified → uses that value
   * - If no `.top()` → uses `pageSize` option (if provided)
   * - If neither → no limit (SharePoint default ~100-5000)
   * 
   * **Note**: If both `.top()` and `pageSize` are specified, `.top()` takes precedence
   * and a warning is logged to the console.
   * 
   * @param queryBuilder - Function to build the query using PnPjs fluent API
   * @param options - Query options (pageSize for pagination)
   * @returns Promise resolving to the array of items
   * 
   * @example Basic query with pageSize
   * ```tsx
   * const { query, items } = useSPFxPnPList('Tasks', { pageSize: 50 });
   * 
   * await query(q => 
   *   q.select('Id', 'Title', 'Status')
   *    .filter("Status eq 'Active'")
   *    .orderBy('Created', false)
   * );
   * ```
   * 
   * @example Query with explicit .top()
   * ```tsx
   * await query(q => 
   *   q.select('Id', 'Title')
   *    .top(100)  // Explicit top takes precedence
   *    .orderBy('Priority', false)
   * );
   * ```
   * 
   * @example Query without pagination
   * ```tsx
   * await query(q => q.select('Id', 'Title'));
   * // Uses SharePoint default limit
   * ```
   */
  query: (
    queryBuilder?: (items: IItems) => IItems,
    options?: { pageSize?: number }
  ) => Promise<T[]>;

  /**
   * Array of list items returned from the last query.
   */
  items: T[];

  /**
   * Indicates if a query is in progress.
   */
  loading: boolean;

  /**
   * Indicates if additional items are being loaded via loadMore().
   */
  loadingMore: boolean;

  /**
   * Error object if an error occurred during any operation.
   */
  error: Error | undefined;

  /**
   * Indicates if the items array is empty.
   */
  isEmpty: boolean;

  /**
   * Indicates if there are more items available to load.
   * Only meaningful when using pagination (pageSize or .top()).
   */
  hasMore: boolean;

  /**
   * Re-executes the last query with the same parameters.
   * Resets pagination state.
   * 
   * @returns Promise resolving when refetch is complete
   * @throws Error if no previous query was executed
   */
  refetch: () => Promise<void>;

  /**
   * Loads the next page of items using the last query.
   * Automatically appends new items to the existing items array.
   * 
   * @returns Promise resolving to the newly loaded items
   * @throws Error if no previous query was executed
   * @throws Error if no pageSize was specified in the previous query
   */
  loadMore: () => Promise<T[]>;

  /**
   * Clears the current error state.
   */
  clearError: () => void;

  /**
   * Retrieves a single list item by ID.
   * 
   * @param id - The ID of the list item to retrieve
   * @returns Promise resolving to the item or undefined if not found
   */
  getById: (id: number) => Promise<T | undefined>;

  /**
   * Creates a new list item.
   * Automatically triggers a refetch after successful creation.
   * 
   * @param item - Partial object containing the fields to set
   * @returns Promise resolving to the ID of the created item
   */
  create: (item: Partial<T>) => Promise<number>;

  /**
   * Updates an existing list item by ID.
   * Automatically triggers a refetch after successful update.
   * 
   * @param id - The ID of the list item to update
   * @param item - Partial object containing the fields to update
   */
  update: (id: number, item: Partial<T>) => Promise<void>;

  /**
   * Deletes a list item by ID.
   * Automatically triggers a refetch after successful deletion.
   * 
   * @param id - The ID of the list item to delete
   */
  remove: (id: number) => Promise<void>;

  /**
   * Creates multiple list items in a single batched request.
   * Automatically triggers a refetch after successful batch creation.
   * 
   * @param items - Array of partial objects to create
   * @returns Promise resolving to an array of created item IDs
   */
  createBatch: (items: Partial<T>[]) => Promise<number[]>;

  /**
   * Updates multiple list items in a single batched request.
   * Automatically triggers a refetch after successful batch update.
   * 
   * @param updates - Array of objects containing id and item fields to update
   */
  updateBatch: (updates: Array<{ id: number; item: Partial<T> }>) => Promise<void>;

  /**
   * Deletes multiple list items in a single batched request.
   * Automatically triggers a refetch after successful batch deletion.
   * 
   * @param ids - Array of item IDs to delete
   */
  removeBatch: (ids: number[]) => Promise<void>;
}

/**
 * Hook for working with SharePoint lists using PnPjs fluent API.
 * Provides query execution with automatic .top() detection, CRUD operations, pagination, and state management.
 * 
 * **Key Features**:
 * - ✅ Native PnPjs fluent API - full type-safe query building
 * - ✅ Smart .top() detection - no conflicts between user .top() and pageSize
 * - ✅ CRUD operations (create, read, update, delete)
 * - ✅ Batch operations for bulk updates
 * - ✅ Pagination with `loadMore()` and `hasMore`
 * - ✅ Automatic refetch after CRUD operations
 * - ✅ Local state management per component instance
 * - ✅ Cross-site support via PnPContextInfo
 * - ✅ ES5 compatibility (IE11 support)
 * 
 * **How .top() Detection Works**:
 * The hook uses a recursive Proxy to detect if `.top()` is called in your queryBuilder:
 * - If `.top()` is specified → uses that value
 * - If no `.top()` but `pageSize` option → adds `.top(pageSize)` automatically
 * - If neither → no limit (SharePoint default ~100-5000)
 * - If both → `.top()` wins, warning logged
 * 
 * @template T - The type of the list item (default: unknown)
 * @param listTitle - The title of the SharePoint list
 * @param options - Optional configuration (pageSize for pagination)
 * @param pnpContext - Optional PnP context for cross-site scenarios
 * @returns Object containing query method, items, loading states, error, and CRUD operations
 * 
 * @example Basic usage with pageSize
 * ```tsx
 * import { useSPFxPnPList } from '@apvee/spfx-react-toolkit';
 * 
 * function TaskList() {
 *   const { query, items, loading, error } = useSPFxPnPList('Tasks', { pageSize: 50 });
 * 
 *   useEffect(() => {
 *     query(q => 
 *       q.select('Id', 'Title', 'Status', 'Priority')
 *        .filter("Status eq 'Active'")
 *        .orderBy('Priority', false)
 *     );
 *   }, [query]);
 * 
 *   if (loading) return <Spinner />;
 *   if (error) return <MessageBar>Error: {error.message}</MessageBar>;
 * 
 *   return (
 *     <ul>
 *       {items.map(task => (
 *         <li key={task.Id}>{task.Title} - {task.Status}</li>
 *       ))}
 *     </ul>
 *   );
 * }
 * ```
 * 
 * @example Type-safe fluent filter
 * ```tsx
 * interface Task {
 *   Id: number;
 *   Title: string;
 *   Status: string;
 *   Priority: number;
 *   DueDate: string;
 * }
 * 
 * function ActiveHighPriorityTasks() {
 *   const { query, items, loading } = useSPFxPnPList<Task>('Tasks', { pageSize: 50 });
 * 
 *   useEffect(() => {
 *     query(q => 
 *       q.select('Id', 'Title', 'Priority', 'DueDate')
 *        .filter(f => f.text("Status").equals("Active")
 *                      .and()
 *                      .number("Priority").greaterThan(3))
 *        .orderBy('DueDate', true)
 *     );
 *   }, [query]);
 * 
 *   return (
 *     <div>
 *       {items.map(task => (
 *         <TaskCard key={task.Id} task={task} />
 *       ))}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Query with explicit .top()
 * ```tsx
 * const { query, items } = useSPFxPnPList<Task>('Tasks');
 * 
 * useEffect(() => {
 *   // User specifies .top() explicitly - takes precedence
 *   query(q => 
 *     q.select('Id', 'Title')
 *      .top(100)  // Hook detects this and uses it
 *      .orderBy('Created', false)
 *   );
 * }, [query]);
 * ```
 * 
 * @example CRUD operations
 * ```tsx
 * function TaskManager() {
 *   const { items, create, update, remove, loading } = useSPFxPnPList<Task>('Tasks');
 * 
 *   const handleCreate = async () => {
 *     try {
 *       const newId = await create({
 *         Title: 'New Task',
 *         Status: 'Active',
 *         Priority: 3
 *       });
 *       console.log('Created task with ID:', newId);
 *       // List automatically refetches
 *     } catch (error) {
 *       console.error('Create failed:', error);
 *     }
 *   };
 * 
 *   const handleUpdate = async (id: number) => {
 *     await update(id, { Status: 'Completed' });
 *     // List automatically refetches
 *   };
 * 
 *   const handleDelete = async (id: number) => {
 *     await remove(id);
 *     // List automatically refetches
 *   };
 * 
 *   return (
 *     <div>
 *       <button onClick={handleCreate}>Create Task</button>
 *       {items.map(task => (
 *         <div key={task.Id}>
 *           <span>{task.Title}</span>
 *           <button onClick={() => handleUpdate(task.Id)}>Complete</button>
 *           <button onClick={() => handleDelete(task.Id)}>Delete</button>
 *         </div>
 *       ))}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Batch operations for bulk updates
 * ```tsx
 * function BulkTaskManager() {
 *   const { createBatch, updateBatch, removeBatch } = useSPFxPnPList<Task>('Tasks');
 * 
 *   const handleBulkCreate = async () => {
 *     const ids = await createBatch([
 *       { Title: 'Task 1', Status: 'Active', Priority: 1 },
 *       { Title: 'Task 2', Status: 'Active', Priority: 2 },
 *       { Title: 'Task 3', Status: 'Active', Priority: 3 }
 *     ]);
 *     console.log('Created task IDs:', ids);
 *   };
 * 
 *   const handleBulkUpdate = async (taskIds: number[]) => {
 *     await updateBatch([
 *       { id: taskIds[0], item: { Status: 'Completed' } },
 *       { id: taskIds[1], item: { Priority: 5 } },
 *       { id: taskIds[2], item: { Status: 'In Progress', Priority: 4 } }
 *     ]);
 *   };
 * 
 *   const handleBulkDelete = async (taskIds: number[]) => {
 *     await removeBatch(taskIds);
 *   };
 * 
 *   return (
 *     <div>
 *       <button onClick={handleBulkCreate}>Create 3 Tasks</button>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Pagination with loadMore
 * ```tsx
 * function InfiniteTaskList() {
 *   const {
 *     query,
 *     items,
 *     loading,
 *     loadingMore,
 *     hasMore,
 *     loadMore,
 *     isEmpty
 *   } = useSPFxPnPList<Task>('Tasks', { pageSize: 50 });
 * 
 *   useEffect(() => {
 *     query(q => 
 *       q.select('Id', 'Title', 'Status')
 *        .orderBy('Created', false)
 *     );
 *   }, [query]);
 * 
 *   if (loading) return <Spinner label="Loading tasks..." />;
 *   if (isEmpty) return <MessageBar>No tasks found</MessageBar>;
 * 
 *   return (
 *     <div>
 *       {items.map(task => (
 *         <TaskCard key={task.Id} task={task} />
 *       ))}
 *       
 *       {hasMore && (
 *         <PrimaryButton
 *           text={loadingMore ? 'Loading...' : 'Load More'}
 *           onClick={loadMore}
 *           disabled={loadingMore}
 *         />
 *       )}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Cross-site usage
 * ```tsx
 * function EmployeeList() {
 *   // Create context for HR site
 *   const hrContext = useSPFxPnPContext({
 *     siteUrl: '/sites/hr'
 *   });
 * 
 *   // Query Employees list from HR site
 *   const { query, items, loading } = useSPFxPnPList<Employee>(
 *     'Employees',
 *     { pageSize: 100 },
 *     hrContext  // Pass context for cross-site query
 *   );
 * 
 *   useEffect(() => {
 *     query(q => 
 *       q.select('Id', 'Name', 'Department', 'Email')
 *        .filter("Department eq 'Engineering'")
 *        .orderBy('Name', true)
 *     );
 *   }, [query]);
 * 
 *   return (
 *     <div>
 *       {items.map(emp => (
 *         <PersonaCard key={emp.Id} employee={emp} />
 *       ))}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Manual refetch with error handling
 * ```tsx
 * function TaskListWithRefresh() {
 *   const { query, items, loading, error, refetch, clearError } = useSPFxPnPList<Task>('Tasks', { pageSize: 50 });
 * 
 *   useEffect(() => {
 *     query(q => q.select('Id', 'Title', 'Status').orderBy('Created', false));
 *   }, [query]);
 * 
 *   return (
 *     <div>
 *       <CommandBar
 *         items={[
 *           {
 *             key: 'refresh',
 *             text: 'Refresh',
 *             iconProps: { iconName: 'Refresh' },
 *             onClick: () => refetch()
 *           }
 *         ]}
 *       />
 * 
 *       {error && (
 *         <MessageBar
 *           messageBarType={MessageBarType.error}
 *           onDismiss={clearError}
 *         >
 *           Error loading tasks: {error.message}
 *         </MessageBar>
 *       )}
 * 
 *       {loading ? (
 *         <Spinner />
 *       ) : (
 *         <DetailsList items={items} />
 *       )}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Conditional loading (autoLoad: false)
 * ```tsx
 * function ConditionalTaskList() {
 *   const [showCompleted, setShowCompleted] = useState(false);
 *   const { items, loading, refetch } = useSPFxPnPList<Task>(
 *     'Tasks',
 *     {
 *       filter: showCompleted ? "Status eq 'Completed'" : "Status eq 'Active'",
 *       autoLoad: false  // Don't load on mount
 *     }
 *   );
 * 
 *   useEffect(() => {
 *     if (showCompleted) {
 *       refetch();  // Manually trigger load
 *     }
 *   }, [showCompleted, refetch]);
 * 
 *   return (
 *     <div>
 *       <Toggle
 *         label="Show Completed"
 *         checked={showCompleted}
 *         onChange={(_, checked) => setShowCompleted(checked || false)}
 *       />
 *       {items.map(task => <TaskCard key={task.Id} task={task} />)}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @remarks
 * **PnPjs Installation**: This hook requires `@pnp/sp` to be installed:
 * ```bash
 * npm install @pnp/sp @pnp/core @pnp/queryable
 * ```
 * 
 * **SharePoint List View Threshold**: Be aware that querying lists with more than 5000 items
 * may cause throttling unless:
 * - Filters use indexed columns
 * - Query results are under 5000 items
 * - Proper pagination is used (top + skip)
 * 
 * **Fluent Filter Requirements**: To use type-safe fluent filters, import types:
 * ```typescript
 * import '@pnp/sp/items';  // Enables fluent filter on items
 * ```
 * 
 * **State Management**: Each hook instance maintains its own local state (items, loading, error).
 * State is not shared between components - this follows the standard React hooks pattern.
 * 
 * **Debounced Refetch**: CRUD operations trigger a debounced refetch (100ms delay) of the
 * current query. A mutation begun under an older query refreshes the current one;
 * replacing the query again before dispatch cancels the scheduled refresh.
 * 
 * @see {@link useSPFxPnPContext} for creating PnP contexts
 * @see {@link PNPContextInfo} for context information
 */
export function useSPFxPnPList<T = unknown>(
  listTitle: string,
  options?: UseSPFxPnPListOptions,
  pnpContext?: PnPContextInfo
): SPFxPnPListInfo<T> {
  return useSPFxPnPListInternal<T>(listTitle, options, pnpContext);
}
