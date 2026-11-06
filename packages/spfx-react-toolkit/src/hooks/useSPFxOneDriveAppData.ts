/**
 * User/application JSON files have separate read/write state and explicit creation policy.
 *
 * @packageDocumentation
 * @docPage hooks/storage/onedrive-app-data
 */
// useSPFxOneDriveAppData.ts
// Hook to manage JSON files in OneDrive approot folder with state management

import { useState, useCallback, useEffect, useRef, useMemo } from 'react';
import { useSPFxMSGraphClient } from './useSPFxMSGraphClient';
import { createSPFxOneDriveAppDataService } from '../services/spfx-onedrive-app-data.service';

/**
 * User/application-scoped OneDrive JSON data, separate read/write state and explicit operations.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxOneDriveAppDataResult, useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const file: SPFxOneDriveAppDataResult<{ theme: string }> = useSPFxOneDriveAppData('settings.json');
 *   console.log(file.data?.theme, file.isNotFound);
 *   return <span>{file.data?.theme ?? 'Unavailable'}</span>;
 * }
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxOneDriveAppDataResult<T> {
  /**
   * Current parsed data or fallback, initially defaultValue. Failed reloads preserve previous data; inspect error/isReady.
   */
  readonly data: T | undefined;

  /** 
   * Loading state for read operations.
   * True while any current-target load remains pending.
   */
  readonly isLoading: boolean;

  /**
   * Error from the latest-started read; cleared when a read starts or the latest write succeeds.
   */
  readonly error: Error | undefined;

  /**
   * True if the file does not exist in OneDrive (404 / itemNotFound).
   * In the legacy signature this is treated as a non-error.
   */
  readonly isNotFound: boolean;

  /** 
   * Loading state for write operations.
   * True while any current-target write remains pending.
   */
  readonly isWriting: boolean;

  /**
   * Error from the latest operation when it was a write; cleared when a new write starts or the target changes. Reads do not clear it.
   */
  readonly writeError: Error | undefined;

  /**
   * Reads Graph metadata and downloads JSON via native fetch without authorization headers/cookies. Missing metadata uses defaultValue; failed reads publish error and resolve. Unavailable clients or empty file names skip the read.
   * @returns Resolves after the read attempt; inspect error and isNotFound.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
   *
   * function Example() {
   *   const file = useSPFxOneDriveAppData<{ theme: string }>('settings.json', { autoFetch: false });
   *   return <button onClick={file.load}>Refresh</button>;
   * }
   * ```
   */
  readonly load: () => Promise<void>;

  /**
   * Creates or replaces JSON file content and publishes successful data when the request is current.
   * @param content - JSON-serializable data to write.
   * @returns Resolves after the write succeeds.
   * @throws Missing/initializing client, empty fileName, serialization or Graph failures.
   */
  readonly write: (content: T) => Promise<void>;

  /**
   * True when !isLoading, no read error and data is defined; fallback/default data can satisfy this without a successful server read, and write status is independent.
   */
  readonly isReady: boolean;
}

/**
 * Optional configuration for useSPFxOneDriveAppData.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxOneDriveAppDataOptions, useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const options: SPFxOneDriveAppDataOptions<{ theme: string }> = { folder: 'dashboard', defaultValue: { theme: 'dark' }, createIfMissing: true };
 *   const file = useSPFxOneDriveAppData('settings.json', options);
 *   return <span>{file.data?.theme ?? 'Unavailable'}</span>;
 * }
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxOneDriveAppDataOptions<T> {
  /**
   * Folder namespace; omitted, undefined or empty uses the app root. Characters other than letters, digits, hyphens and underscores become hyphens. fileName is not sanitized.
   */
  folder?: string;

  /** Whether to load on mount; omitted or undefined uses true.
   * @defaultValue true
   */
  autoFetch?: boolean;

  /**
   * Initial/fallback data when a file is confirmed missing; omitted or undefined leaves data undefined. Changing it alone does not reset current data; future missing reads use the latest value.
   */
  defaultValue?: T;

  /**
   * Omitted or undefined uses false. After a current confirmed missing read, writes a defined defaultValue once per load attempt; failed downloads do not trigger creation.
   * @defaultValue false
   */
  createIfMissing?: boolean;
}

/**
 * Manages JSON content in the authenticated user's OneDrive app root. Requires Graph permissions for the requested read/write operations and native fetch for preauthenticated downloads. T is an expected shape without runtime validation. Auto-load defaults to true; target changes reset state and obsolete completions cannot publish it. Reads and writes have separate pending counters but share latest-started result ownership, so a later write prevents an earlier read from publishing stale data. Retained callbacks keep their original remote target; dispatched requests are not cancelled. Read failures publish error and resolve; writes reject and publish writeError. Recognized missing metadata returns fallback data without an error.
 * @typeParam T - Expected JSON data shape.
 * @param fileName - File name, passed into the Graph path without sanitization.
 * @param options - Folder, auto-load, fallback and creation settings; omitted or undefined uses app root, auto-load true, undefined fallback and creation false.
 * @returns Read/write state and explicit load/write methods.
 * @throws When called outside an SPFx provider.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 *
 * function Settings() {
 *   const file = useSPFxOneDriveAppData<{ theme: string }>('settings.json', {
 *     folder: 'dashboard',
 *     autoFetch: false,
 *     defaultValue: { theme: 'dark' }
 *   });
 *   return (
 *     <button onClick={file.load} disabled={file.isLoading}>
 *       Load settings
 *     </button>
 *   );
 * }
 * ```
 *
 * @useWhen JSON settings belong to the authenticated user and application in OneDrive.
 *
 *
 * @example Create preferences only for a missing app file
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 * interface UserPrefs { favorites: string[]; notifications: boolean }
 * function UserPreferences() {
 *   const prefs = useSPFxOneDriveAppData<UserPrefs>('prefs.json', {
 *     folder: 'user-settings', defaultValue: { favorites: [], notifications: true }, createIfMissing: true,
 *   });
 *   const add = async () => {
 *     if (prefs.data) await prefs.write({ ...prefs.data, favorites: [...prefs.data.favorites, 'report'] });
 *   };
 *   return <div><p>Favorites: {prefs.data?.favorites.length ?? 0}</p>
 *     <button disabled={!prefs.isReady || prefs.isWriting} onClick={() => void add().catch(console.error)}>Add favorite</button>
 *     {(prefs.error ?? prefs.writeError) && <p role="alert">{(prefs.error ?? prefs.writeError)?.message}</p>}
 *   </div>;
 * }
 * ```
 *
 *
 * @example Load an app-folder cache on demand
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 * function LazyCache() {
 *   const cache = useSPFxOneDriveAppData<unknown>('cache.json', { folder: 'my-app-v2', autoFetch: false });
 *   return <div><button disabled={cache.isLoading} onClick={() => void cache.load().catch(console.error)}>Load cache</button>
 *     {cache.error && <p role="alert">{cache.error.message}</p>}
 *     {cache.isReady && <pre>{JSON.stringify(cache.data, undefined, 2)}</pre>}
 *   </div>;
 * }
 * ```
 *
 *
 * @example Add toggle and remove persisted todos
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 * interface TodoData { items: { id: string; text: string; done: boolean }[] }
 * function TodoApp() {
 *   const store = useSPFxOneDriveAppData<TodoData>('todos.json', {
 *     folder: 'todo-app', defaultValue: { items: [] }, createIfMissing: true,
 *   });
 *   const [text, setText] = React.useState('');
 *   const add = async () => { await store.write({ items: [...(store.data?.items ?? []), { id: crypto.randomUUID(), text, done: false }] }); };
 *   const toggle = async (id: string) => { await store.write({ items: (store.data?.items ?? []).map(item => item.id === id ? { ...item, done: !item.done } : item) }); };
 *   const remove = async (id: string) => { await store.write({ items: (store.data?.items ?? []).filter(item => item.id !== id) }); };
 *   return <div><input aria-label="New todo" value={text} onChange={event => setText(event.target.value)} />
 *     <button disabled={!store.isReady || store.isWriting} onClick={() => void add().catch(console.error)}>Add</button>
 *     <ul>{store.data?.items.map(item => <li key={item.id}>{item.text}
 *       <button disabled={store.isWriting} onClick={() => void toggle(item.id).catch(console.error)}>{item.done ? 'Reopen' : 'Complete'}</button>
 *       <button disabled={store.isWriting} onClick={() => void remove(item.id).catch(console.error)}>Delete</button>
 *     </li>)}</ul>
 *     {(store.error ?? store.writeError) && <p role="alert">{(store.error ?? store.writeError)?.message}</p>}
 *   </div>;
 * }
 * ```
 *
 *
 * @example Edit app-folder configuration
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 * function ConfigPanel() {
 *   const config = useSPFxOneDriveAppData<{ theme: 'light' | 'dark'; language: string }>('config.json');
 *   const change = async (dark: boolean) => { if (config.data) await config.write({ ...config.data, theme: dark ? 'dark' : 'light' }); };
 *   return <div><label>Dark mode<input type="checkbox" checked={config.data?.theme === 'dark'} disabled={!config.isReady || config.isWriting}
 *     onChange={event => void change(event.target.checked).catch(console.error)} /></label>
 *     {(config.error ?? config.writeError) && <p role="alert">{(config.error ?? config.writeError)?.message}</p>}
 *   </div>;
 * }
 * ```
 *
 * @remarks
 * Mount below the matching SPFx provider with Fluent peers and approved app-folder grants. Start gates all Graph initialization. This scenario disables autoFetch and createIfMissing and supplies no defaultValue. Only explicit Write replaces the fixed file; it does not create a folder. Reset preserves the operational child, original hook callbacks and pending read/write locks, restores local drafts and hides stale output without clearing hook state, cancelling I/O or rolling back accepted writes. T remains an expected JSON shape, not runtime validation.
 *
 * @scenario Read and explicitly replace OneDrive app data
 * ```tsx
 * import * as React from 'react';
 * import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Field, Input, Label, MessageBar, MessageBarBody, Select, makeStyles } from '@fluentui/react-components';
 * import { useSPFxMSGraphClient, useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * type DemoProps = { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number };
 * interface OneDriveDemoData { readonly message: string; readonly counter: number; readonly timestamp: number; }
 * const useStyles = makeStyles({ diagnostics: { minWidth: 0, maxWidth: '100%', overflowX: 'auto' }, json: { whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' } });
 * function ActiveOneDrive({ render, resetRevision = 0 }: DemoProps): React.ReactElement {
 *   const classes = useStyles();
 *   const options = React.useMemo(() => ({ autoFetch: false, createIfMissing: false, folder: 'spfx-react-toolkit-demo' }), []);
 *   const graph = useSPFxMSGraphClient();
 *   const appData = useSPFxOneDriveAppData<OneDriveDemoData>('test-data.json', options);
 *   const [operation, setOperation] = React.useState<'read' | 'write'>('read');
 *   const [draft, setDraft] = React.useState({ message: 'Project update', revision: 0 });
 *   const message = draft.message;
 *   const live = React.useRef({ graph, appData, message });
 *   live.current = { graph, appData, message };
 *   const identity = React.useRef({ client: graph.client, load: appData.load, write: appData.write, revision: resetRevision, operation, message, draftRevision: draft.revision });
 *   if (identity.current.client !== graph.client || identity.current.load !== appData.load || identity.current.write !== appData.write || identity.current.revision !== resetRevision || identity.current.operation !== operation || identity.current.draftRevision !== draft.revision) {
 *     identity.current = { client: graph.client, load: appData.load, write: appData.write, revision: resetRevision, operation, message, draftRevision: draft.revision };
 *   }
 *   const boundary = identity.current;
 *   const mounted = React.useRef(true);
 *   React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
 *   const sequence = React.useRef(0);
 *   const pending = React.useRef<{ load: typeof appData.load; write: typeof appData.write; request: number }>();
 *   const [executed, setExecuted] = React.useState<{ boundary: typeof boundary; request: number; completed: boolean; error?: Error }>();
 *   const observedRevision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (observedRevision.current !== resetRevision) {
 *       observedRevision.current = resetRevision;
 *       setOperation('read'); setDraft(previous => ({ message: 'Project update', revision: previous.revision + 1 })); setExecuted(undefined);
 *     }
 *   }, [resetRevision]);
 *   const run = React.useCallback(async (): Promise<void> => {
 *     const active = live.current;
 *     if (!mounted.current || identity.current !== boundary || active.message !== boundary.message || !active.graph.isReady || !active.graph.client || active.appData.isLoading || active.appData.isWriting || (pending.current?.load === boundary.load && pending.current.write === boundary.write) || (boundary.operation === 'write' && boundary.message === '')) return;
 *     const request = ++sequence.current;
 *     pending.current = { load: boundary.load, write: boundary.write, request };
 *     const isCurrent = (): boolean => mounted.current && identity.current === boundary && sequence.current === request;
 *     setExecuted({ boundary, request, completed: false });
 *     try {
 *       if (boundary.operation === 'read') await boundary.load();
 *       else {
 *         await boundary.write({ message: boundary.message, counter: (active.appData.data?.counter ?? 0) + 1, timestamp: Date.now() });
 *       }
 *       if (isCurrent()) {
 *         // User edits retire the draft revision; this owned clear preserves its completed result.
 *         if (boundary.operation === 'write') setDraft(previous => ({ ...previous, message: '' }));
 *         setExecuted({ boundary, request, completed: true });
 *       }
 *     } catch (failure) {
 *       if (isCurrent()) setExecuted({ boundary, request, completed: true, error: failure instanceof Error ? failure : new Error(String(failure)) });
 *     } finally {
 *       if (pending.current?.request === request) pending.current = undefined;
 *     }
 *   }, [boundary]);
 *   const available = graph.isReady && graph.client !== undefined;
 *   const busy = appData.isLoading || appData.isWriting;
 *   const current = executed?.boundary === boundary && executed.request === sequence.current;
 *   const error = current ? executed?.error ?? (operation === 'read' ? appData.error : appData.writeError) : undefined;
 *   const completed = current && executed?.completed && !busy;
 *   // Legacy JSON null remains present; optional field reads do not validate stored shapes.
 *   const regions: DemoRegions = {
 *     controls: <><Field label="OneDrive operation"><Select aria-label="OneDrive operation" disabled={appData.isWriting} value={operation} onChange={(_, data) => { if (data.value === 'read' || data.value === 'write') setOperation(data.value); }}><option value="read">Read</option><option value="write">Write</option></Select></Field>{operation === 'write' && <Field label="Message"><Input aria-label="Message" disabled={appData.isWriting} value={message} onChange={(_, data) => setDraft(previous => ({ message: data.value, revision: previous.revision + 1 }))} /></Field>}</>,
 *     commands: <><Button appearance="primary" disabled={!available || busy || (operation === 'write' && message === '')} onClick={() => { run().catch(() => undefined); }}>Run</Button></>,
 *     status: graph.initError ? <MessageBar intent="error"><MessageBarBody>{graph.initError.message}</MessageBarBody></MessageBar> : error ? <MessageBar intent="error"><MessageBarBody>{error.message}</MessageBarBody></MessageBar> : undefined,
 *     children: !available ? <Label>Microsoft Graph unavailable. Wait for client initialization.</Label> : busy ? <Label>Loading...</Label> : !completed ? <Label>Run the selected operation.</Label> : error ? <Label>Operation not completed.</Label> : appData.isNotFound ? <Label>File not found.</Label> : appData.data === undefined ? <Label>No data returned.</Label> : <dl><dt>Saved message</dt><dd>{appData.data?.message === '' ? '(empty)' : appData.data?.message}</dd><dt>Counter</dt><dd>{appData.data ? String(appData.data.counter) : undefined}</dd></dl>,
 *     diagnostics: <><Label>approot:/spfx-react-toolkit-demo/test-data.json</Label><dl><dt>Ready</dt><dd>{completed ? String(appData.isReady) : 'Not requested'}</dd><dt>Not Found</dt><dd>{completed ? String(appData.isNotFound) : 'Not requested'}</dd></dl>{completed && !error && appData.data !== undefined && <Accordion collapsible className={classes.diagnostics}><AccordionItem value="data"><AccordionHeader>OneDrive data</AccordionHeader><AccordionPanel><pre className={classes.json}>{JSON.stringify(appData.data, undefined, 2)}</pre></AccordionPanel></AccordionItem></Accordion>}<MessageBar intent="info"><MessageBarBody>Reading a missing file does not create it. Only explicit Write replaces this file. Reset restores local drafts and hides publication; it preserves the active hook and genuine pending locks, does not cancel requests and cannot undo a remote write. Generic JSON types do not validate stored data.</MessageBarBody></MessageBar></>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export function Demo({ render, resetRevision = 0 }: DemoProps = {}): React.ReactElement {
 *   const [started, setStarted] = React.useState(false);
 *   if (started) return <ActiveOneDrive render={render} resetRevision={resetRevision} />;
 *   const regions: DemoRegions = { commands: <><Button appearance="primary" onClick={() => setStarted(true)}>Start</Button></>, children: <p>Start initializes Microsoft Graph under the matching SPFx provider. Then choose Read or explicitly Write the fixed app-folder file.</p> };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Start initializes Graph without reading or writing. Read missing data never creates a file; explicit Write replaces the fixed app-folder message and increments its counter. Pending writes lock inputs. Reset restores local drafts and hides old publication while preserving the hook and real pending locks; it cannot undo remote writes.' };
 * ```
 */
export function useSPFxOneDriveAppData<T = unknown>(
  fileName: string,
  options?: SPFxOneDriveAppDataOptions<T>
): SPFxOneDriveAppDataResult<T> {
  const {
    client,
    isReady: isClientReady,
    isInitializing: isClientInitializing,
    initError: clientInitError
  } = useSPFxMSGraphClient();

  // ═══════════════════════════════════════════════════════════════════════════
  // OPTIONS (stabilized with useMemo to prevent unnecessary re-renders)
  // ═══════════════════════════════════════════════════════════════════════════

  const resolvedOptions = useMemo(() => options ?? {}, [options]);

  // Extract stable primitive values for dependency arrays
  const folder = resolvedOptions.folder;
  const shouldAutoFetch = resolvedOptions.autoFetch ?? true;
  const defaultValue = resolvedOptions.defaultValue;
  const defaultValueRef = useRef(defaultValue);
  defaultValueRef.current = defaultValue;
  const createIfMissing = resolvedOptions.createIfMissing ?? false;

  const oneDriveAppDataService = useMemo(() => (
    client ? createSPFxOneDriveAppDataService(client) : undefined
  ), [client]);

  // ═══════════════════════════════════════════════════════════════════════════
  // STATE MANAGEMENT
  // ═══════════════════════════════════════════════════════════════════════════

  const [data, setData] = useState<T | undefined>(defaultValue);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error | undefined>(undefined);
  const [isWriting, setIsWriting] = useState<boolean>(false);
  const [writeError, setWriteError] = useState<Error | undefined>(undefined);
  const [isNotFound, setIsNotFound] = useState<boolean>(false);

  // ═══════════════════════════════════════════════════════════════════════════
  // REFS (for cleanup and stable references)
  // ═══════════════════════════════════════════════════════════════════════════

  // Track component mounted state to prevent memory leaks
  const isMountedRef = useRef<boolean>(true);
  useEffect(() => {
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  const createAttemptedRef = useRef<boolean>(false);
  const identityRef = useRef({ service: oneDriveAppDataService, fileName, folder,
    latest: 0, reads: 0, writes: 0, missingRequest: 0 });
  if (identityRef.current.service !== oneDriveAppDataService ||
      identityRef.current.fileName !== fileName || identityRef.current.folder !== folder) {
    identityRef.current = { service: oneDriveAppDataService, fileName, folder,
      latest: 0, reads: 0, writes: 0, missingRequest: 0 };
  }
  useEffect(() => {
    setData(defaultValueRef.current);
    setError(undefined);
    setWriteError(undefined);
    setIsLoading(false);
    setIsWriting(false);
    setIsNotFound(false);
    createAttemptedRef.current = false;
  }, [oneDriveAppDataService, fileName, folder]);

  const currentIdentity = identityRef.current;

  // ═══════════════════════════════════════════════════════════════════════════
  // WRITE CALLBACK (defined first, no dependency on load)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Write data to OneDrive file
   * Creates file if it doesn't exist, updates if it does (upsert)
   * Updates isWriting and writeError states
   */
  const write = useCallback(async (content: T): Promise<void> => {
    const service = oneDriveAppDataService;

    if (!client || !service) {
      if (isClientInitializing) {
        throw new Error('Graph client is still initializing. Please wait and try again.');
      }
      if (clientInitError) {
        throw new Error(`Graph client initialization failed: ${clientInitError.message}`);
      }
      throw new Error('Graph client not available. Cannot write file.');
    }

    if (!fileName) {
      throw new Error('fileName is required. Cannot write file.');
    }

    const identity = currentIdentity;
    const request = ++identity.latest;
    const isIdentityCurrent = (): boolean => isMountedRef.current && identityRef.current === identity &&
      identity.service === service && identity.fileName === fileName && identity.folder === folder;
    identity.writes++;
    identity.missingRequest = 0;
    if (isIdentityCurrent()) { setIsWriting(true); setWriteError(undefined); }

    try {
      await service.write<T>(fileName, content, folder);

      if (isIdentityCurrent() && request === identity.latest) {
        // Update local data to reflect successful write
        setData(content);
        setIsNotFound(false);
        // Clear read error if write succeeds (fresh state)
        setError(undefined);
      }
    } catch (err) {
      if (isIdentityCurrent() && request === identity.latest) {
        const writeErr = err instanceof Error ? err : new Error(String(err));
        setWriteError(writeErr);
        console.error('Failed to write file to OneDrive:', writeErr);
      }
      // Re-throw to allow caller to handle
      throw err;
    } finally {
      identity.writes--;
      if (isIdentityCurrent()) setIsWriting(identity.writes > 0);
    }
  }, [client, oneDriveAppDataService, fileName, folder, isClientInitializing, clientInitError, currentIdentity]);

  // ═══════════════════════════════════════════════════════════════════════════
  // LOAD CALLBACK (NO dependency on write - uses effect for createIfMissing)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Load file from OneDrive
   * Updates data, isLoading, and error states
   * Does NOT call write directly - createIfMissing is handled by separate effect
   */
  const load = useCallback(async (): Promise<void> => {
    const service = oneDriveAppDataService;

    if (!client || !service) {
      if (isClientInitializing) {
        console.info('Graph client is still initializing. Skipping load - will auto-retry when ready.');
        return;
      }
      if (clientInitError) {
        console.error('Graph client initialization failed:', clientInitError.message);
        return;
      }
      console.warn('Graph client not available. Skipping load.');
      return;
    }

    if (!fileName) {
      console.warn('fileName is required. Skipping load.');
      return;
    }

    // Reset createAttempted flag when load is called (fresh attempt)
    const identity = currentIdentity;
    if (identityRef.current === identity) createAttemptedRef.current = false;
    const request = ++identity.latest;
    const fallbackValue = defaultValueRef.current;
    const isIdentityCurrent = (): boolean => isMountedRef.current && identityRef.current === identity &&
      identity.service === service && identity.fileName === fileName && identity.folder === folder;
    identity.reads++;
    identity.missingRequest = 0;
    if (isIdentityCurrent()) {
      setIsLoading(true);
      setError(undefined);
      setIsNotFound(false);
    }

    try {
      const readResult = await service.read<T>(fileName, folder);

      if (isIdentityCurrent() && request === identity.latest) {
        if (readResult.isNotFound) {
          identity.missingRequest = request;
          // Missing file is treated as a non-error.
          // Set data to defaultValue if provided, otherwise undefined
          if (fallbackValue !== undefined) {
            setData(fallbackValue);
          } else {
            setData(undefined);
          }
          setIsNotFound(true);
          setError(undefined);
          console.info('OneDrive file not found. isNotFound=true');
          // NOTE: createIfMissing is handled by separate useEffect
          return;
        }

        setData(readResult.data);
        setIsNotFound(false);
      }
    } catch (err) {
      if (isIdentityCurrent() && request === identity.latest) {
        setIsNotFound(false);
        const loadError = err instanceof Error ? err : new Error(String(err));
        setError(loadError);
        console.error('Failed to load file from OneDrive:', loadError);
      }
    } finally {
      identity.reads--;
      if (isIdentityCurrent()) setIsLoading(identity.reads > 0);
    }
  }, [client, oneDriveAppDataService, fileName, folder, isClientInitializing, clientInitError, currentIdentity]); // ← NO write, NO createIfMissing

  // ═══════════════════════════════════════════════════════════════════════════
  // EFFECTS
  // ═══════════════════════════════════════════════════════════════════════════

  // Auto-fetch on mount if enabled (wait for client to be ready)
  useEffect(() => {
    if (shouldAutoFetch && isClientReady && fileName) {
      load().catch(() => {
        // Error already handled in load() function
      });
    }
  }, [shouldAutoFetch, isClientReady, fileName, load]);

  // Separate effect for createIfMissing - reacts to isNotFound state
  // This breaks the circular dependency: load → write
  useEffect(() => {
    // Guard conditions:
    // 1. File must be not found
    // 2. createIfMissing must be enabled
    // 3. defaultValue must be provided
    // 4. Must not be currently writing (prevent double-write)
    // 5. Must not have already attempted create (prevent infinite loop)
    // 6. Must not be currently loading (wait for load to complete)
    if (
      isNotFound &&
      identityRef.current.missingRequest > 0 &&
      identityRef.current.missingRequest === identityRef.current.latest &&
      createIfMissing &&
      defaultValue !== undefined &&
      !isWriting &&
      !createAttemptedRef.current &&
      !isLoading
    ) {
      createAttemptedRef.current = true;
      write(defaultValue).catch((writeErr) => {
        // write() already updates writeError state
        console.error('Failed to create missing file in OneDrive:', writeErr);
      });
    }
  }, [isNotFound, createIfMissing, defaultValue, isWriting, isLoading, write]);

  // ═══════════════════════════════════════════════════════════════════════════
  // COMPUTED STATE & RETURN
  // ═══════════════════════════════════════════════════════════════════════════

  // Computed state: ready when data loaded successfully
  const isReady = !isLoading && !error && data !== undefined;

  return useMemo(() => ({
    data,
    isLoading,
    error,
    isWriting,
    writeError,
    isNotFound,
    load,
    write,
    isReady,
  }), [data, isLoading, error, isWriting, writeError, isNotFound, load, write, isReady]);
}
