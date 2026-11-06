/**
 * Browser storage uses an instance namespace and best-effort JSON persistence.
 *
 * @packageDocumentation
 * @docPage hooks/storage/local-storage
 */
// useSPFxStorage.ts
// Hooks for persisted storage scoped to SPFx instance

import { useMemo, useCallback, useEffect, useRef, useState } from 'react';
import { createScopedSPFxStorageKey } from '../helpers/spfx-storage.helpers';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';

/**
 * Instance-scoped browser JSON state and best-effort persistence operations.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxStorageHook, useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const storage: SPFxStorageHook<number> = useSPFxLocalStorage('pageSize', 25);
 *   const onIncrease = () => storage.setValue(value => value + 5);
 *   return <button onClick={()=>void onIncrease()}>Run operation</button>;
 * }
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 *
 * @docPage hooks/storage/local-storage
 */
export interface SPFxStorageHook<T> {
  /** Current value */
  readonly value: T;
  
  /**
   * Updates React state and attempts JSON persistence; storage/serialization failures are ignored. Function values are treated as updaters.
   * @param value - Next value or pure updater receiving current state.
   * @returns No value.
   */
  readonly setValue: (value: T | ((prev: T) => T)) => void;
  
  /**
   * Attempts to delete the scoped storage entry and resets state to the latest defaultValue. Persistence failures are ignored.
   * @returns No value.
   */
  readonly remove: () => void;
}

type StorageKind = 'local' | 'session';

type StorageValueUpdater<T> = T | ((prev: T) => T);

function getBrowserStorage(kind: StorageKind): Storage | undefined {
  try {
    if (kind === 'local' && typeof localStorage !== 'undefined') {
      return localStorage;
    }

    if (kind === 'session' && typeof sessionStorage !== 'undefined') {
      return sessionStorage;
    }
  } catch {
    return undefined;
  }

  return undefined;
}

function readStorageValue<T>(
  storage: Storage | undefined,
  key: string,
  defaultValue: T
): T {
  if (!storage) {
    return defaultValue;
  }

  try {
    const item = storage.getItem(key);
    if (item === null) {
      return defaultValue;
    }

    return JSON.parse(item) as T;
  } catch {
    return defaultValue;
  }
}

function writeStorageValue<T>(
  storage: Storage | undefined,
  key: string,
  value: T
): void {
  if (!storage) {
    return;
  }

  try {
    const item = JSON.stringify(value);
    if (item === undefined) {
      storage.removeItem(key);
      return;
    }

    storage.setItem(key, item);
  } catch {
    // Storage can fail in private browsing, quota limits, or blocked contexts.
  }
}

function removeStorageValue(storage: Storage | undefined, key: string): void {
  if (!storage) {
    return;
  }

  try {
    storage.removeItem(key);
  } catch {
    // Best-effort persistence API.
  }
}

function useSPFxBrowserStorage<T>(
  kind: StorageKind,
  key: string,
  defaultValue: T
): SPFxStorageHook<T> {
  const { id: instanceId } = useSPFxInstanceInfo();
  const scopedKey = useMemo(
    () => createScopedSPFxStorageKey(instanceId, key),
    [instanceId, key]
  );

  // Defaults initialize the current key; changing an inline default must not reset it.
  // Removal and external deletion still use the latest caller-provided fallback.
  const defaultValueRef = useRef(defaultValue);
  defaultValueRef.current = defaultValue;

  const [value, setStoredValue] = useState<T>(() => (
    readStorageValue(getBrowserStorage(kind), scopedKey, defaultValue)
  ));

  useEffect(() => {
    setStoredValue(readStorageValue(getBrowserStorage(kind), scopedKey, defaultValueRef.current));
  }, [kind, scopedKey]);

  const setValue = useCallback((nextValue: StorageValueUpdater<T>): void => {
    setStoredValue(previous => {
      const resolved = typeof nextValue === 'function'
        ? (nextValue as (prev: T) => T)(previous)
        : nextValue;

      writeStorageValue(getBrowserStorage(kind), scopedKey, resolved);
      return resolved;
    });
  }, [kind, scopedKey]);

  const remove = useCallback((): void => {
    removeStorageValue(getBrowserStorage(kind), scopedKey);
    setStoredValue(defaultValueRef.current);
  }, [kind, scopedKey]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    const storage = getBrowserStorage(kind);
    if (!storage) {
      return;
    }

    const handleStorage = (event: StorageEvent): void => {
      if (event.storageArea !== storage || event.key !== scopedKey) {
        return;
      }

      setStoredValue(readStorageValue(storage, scopedKey, defaultValueRef.current));
    };

    window.addEventListener('storage', handleStorage);

    return () => {
      window.removeEventListener('storage', handleStorage);
    };
  }, [kind, scopedKey]);

  return useMemo(() => ({
    value,
    setValue,
    remove,
  }), [value, setValue, remove]);
}

/**
 * Hook to use localStorage scoped to SPFx instance
 *
 * The storage key is automatically scoped to the SPFx instance ID,
 * ensuring isolation between different web parts/extensions.
 *
 * Data persists across page reloads and sessions.
 *
 * Use for:
 * - User preferences
 * - Form drafts
 * - Long-lived cache
 * - Settings
 *
 * The key is `spfx:<instanceId>:<key>`, an instance namespace rather than an authorization boundary. The default seeds React state but is not written until setValue. Key/instance changes reload that key; changing only the default preserves current state. remove and exact-key browser deletion events use the latest default. Storage/JSON failures are ignored; local state can change even when persistence fails. T supplies no runtime validation, and only JSON-compatible values reliably round-trip. Same-document hook instances do not synchronize via storage events; clear events with no matching key are ignored.
 *
 * @param key - Storage key (will be prefixed with instance ID)
 * @param defaultValue - Fallback when initializing a key or removing its value. Changing it alone preserves the current value; remove uses the latest fallback.
 * @returns Storage hook with value, setValue, and remove
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';
 *
 * function MyComponent() {
 *   const { value, setValue } = useSPFxLocalStorage('view-mode', 'grid');
 *
 *   return (
 *     <div>
 *       <p>View: {value}</p>
 *       <button onClick={() => setValue('list')}>List View</button>
 *       <button onClick={() => setValue('grid')}>Grid View</button>
 *     </div>
 *   );
 * }
 * ```
 * @throws When called outside an SPFx provider.
 * @template T - Expected application data shape; no runtime validation is performed.
 *
 * @docPage hooks/storage/local-storage
 *
 * @useWhen JSON preferences should persist in this browser across sessions for this SPFx instance.
 *
 * @avoidWhen data must be shared across browsers or validated on the server; use a server-backed store.
 *
 *
 * @example Edit local preferences
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';
 *
 * interface UserPreferences {
 *   theme: 'light' | 'dark';
 *   itemsPerPage: number;
 *   showSidebar: boolean;
 * }
 *
 * function SettingsPanel() {
 *   const { value: preferences, setValue: setPreferences } = useSPFxLocalStorage<UserPreferences>(
 *     'userPrefs',
 *     { theme: 'light', itemsPerPage: 10, showSidebar: true }
 *   );
 *
 *   const updateTheme = (theme: 'light' | 'dark') => {
 *     setPreferences(prev => ({ ...prev, theme }));
 *   };
 *
 *   const updateItemsPerPage = (itemsPerPage: number) => {
 *     setPreferences(prev => ({ ...prev, itemsPerPage }));
 *   };
 *
 *   return (
 *     <div className="settings">
 *       <label>
 *         Theme:
 *         <select
 *           value={preferences.theme}
 *           onChange={e => updateTheme(e.target.value === 'dark' ? 'dark' : 'light')}
 *         >
 *           <option value="light">Light</option>
 *           <option value="dark">Dark</option>
 *         </select>
 *       </label>
 *
 *       <label>
 *         Items per page:
 *         <input
 *           type="number"
 *           value={preferences.itemsPerPage}
 *           onChange={e => updateItemsPerPage(Number(e.target.value))}
 *           min={5}
 *           max={50}
 *         />
 *       </label>
 *     </div>
 *   );
 * }
 * ```
 *
 *
 * @example Keep and clear the last ten searches
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';
 *
 * function SearchWithHistory() {
 *   const { value: history, setValue: setHistory } = useSPFxLocalStorage<string[]>('searchHistory', []);
 *   const [query, setQuery] = React.useState('');
 *
 *   const handleSearch = () => {
 *     if (query && !history.includes(query)) {
 *       setHistory(prev => [query, ...prev.slice(0, 9)]); // Keep last 10
 *     }
 *     // Perform search...
 *   };
 *
 *   const clearHistory = () => setHistory([]);
 *
 *   return (
 *     <div>
 *       <input
 *         value={query}
 *         onChange={e => setQuery(e.target.value)}
 *         placeholder="Search..."
 *         list="search-history"
 *       />
 *       <datalist id="search-history">
 *         {history.map(term => (
 *           <option key={term} value={term} />
 *         ))}
 *       </datalist>
 *       <button onClick={handleSearch}>Search</button>
 *       <button onClick={clearHistory}>Clear History</button>
 *     </div>
 *   );
 * }
 * ```
 *
 * @scenario Save and clear an instance scoped local value
 * ```tsx
 * // Requires the matching SPFx provider. Only named commands write; browser JSON persistence is best effort. The string generic does not validate stored JSON. Reset keeps current hook state, rather than rereading persistence; reentry initializes from the scoped key.
 * import * as React from 'react';
 * import { Button, Input } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   children?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const { spfxContext } = useSPFxContext();
 *   const { value, setValue, remove } = useSPFxLocalStorage('demo-local-key', '');
 *   const [draft, setDraft] = React.useState('');
 *   const [message, setMessage] = React.useState<string>();
 *   const lifetime = React.useMemo(() => ({ active: true, timer: undefined as number | undefined }),
 *     [resetRevision, spfxContext, setValue, remove]);
 *   React.useLayoutEffect(() => {
 *     setDraft('');
 *     setMessage(undefined);
 *     return () => {
 *       lifetime.active = false;
 *       if (lifetime.timer !== undefined) window.clearTimeout(lifetime.timer);
 *     };
 *   }, [lifetime]);
 *   const save = (): void => {
 *     if (!lifetime.active || !draft) return;
 *     setValue(draft);
 *     setDraft('');
 *     setMessage('Local value updated; persistence attempted.');
 *     if (lifetime.timer !== undefined) window.clearTimeout(lifetime.timer);
 *     const timer = window.setTimeout(() => {
 *       if (!lifetime.active || lifetime.timer !== timer) return;
 *       lifetime.timer = undefined;
 *       setMessage(undefined);
 *     }, 3000);
 *     lifetime.timer = timer;
 *   };
 *   const clear = (): void => { if (lifetime.active) remove(); };
 *   return render({
 *     controls: <Input aria-label="Local value" value={draft} onChange={(_, data) => setDraft(data.value)} />,
 *     commands: <><Button disabled={!draft} onClick={save}>Save Local</Button><Button onClick={clear}>Clear Local</Button></>,
 *     status: message ? <p role="status">{message}</p> : undefined,
 *     children: <output>Current value: {String(value ?? '')}</output>,
 *   });
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Save Local updates the current instance value and attempts browser persistence; Clear Local removes its scoped key and restores the empty fallback without new feedback. Reset clears only input/feedback and keeps the current hook value, without rereading browser storage.' };
 * ```
 */
export function useSPFxLocalStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T> {
  return useSPFxBrowserStorage('local', key, defaultValue);
}

/**
 * Hook to use sessionStorage scoped to SPFx instance
 *
 * The storage key is automatically scoped to the SPFx instance ID,
 * ensuring isolation between different web parts/extensions.
 *
 * Data persists only for the current browser session/tab.
 *
 * Use for:
 * - Temporary state
 * - Session-specific cache
 * - Tab-specific settings
 * - Wizard state
 *
 * Shares local storage's JSON/default/reset and best-effort failure rules. The key uses the same instance namespace; browser per-origin/per-tab session rules determine sharing. Same-document hook instances do not synchronize automatically, and an identical key in another tab is not a shared persistence channel.
 *
 * @param key - Storage key (will be prefixed with instance ID)
 * @param defaultValue - Fallback when initializing a key or removing its value. Changing it alone preserves the current value; remove uses the latest fallback.
 * @returns Storage hook with value, setValue, and remove
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';
 *
 * function WizardComponent() {
 *   const { value: step, setValue: setStep } = useSPFxSessionStorage('wizard-step', 1);
 *
 *   return (
 *     <div>
 *       <p>Step: {step}</p>
 *       <button onClick={() => setStep(s => s + 1)}>Next</button>
 *     </div>
 *   );
 * }
 * ```
 * @throws When called outside an SPFx provider.
 * @template T - Expected application data shape; no runtime validation is performed.
 *
 * @docPage hooks/storage/session-storage
 *
 * @useWhen JSON workflow state should last only within this browser tab/session for this instance.
 *
 * @avoidWhen data must survive the tab/session or be shared with another browser; choose persistent or server storage.
 *
 *
 * @example Save a session draft before clearing it
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';
 *
 * interface FormData {
 *   title: string;
 *   description: string;
 *   category: string;
 * }
 *
 * function FormWithDraft({ save }: { save: (draft: FormData) => Promise<void> }) {
 *   const { value: draft, setValue: setDraft } = useSPFxSessionStorage<FormData>('formDraft', {
 *     title: '',
 *     description: '',
 *     category: ''
 *   });
 *
 *   const handleChange = (field: keyof FormData, value: string) => {
 *     setDraft(prev => ({ ...prev, [field]: value }));
 *   };
 *
 *   const handleSubmit = async () => {
 *     await save(draft);
 *     setDraft({ title: '', description: '', category: '' }); // Clear draft
 *   };
 *
 *   return (
 *     <form onSubmit={e => { e.preventDefault(); void handleSubmit().catch(console.error); }}>
 *       <input
 *         value={draft.title}
 *         onChange={e => handleChange('title', e.target.value)}
 *         placeholder="Title"
 *       />
 *       <textarea
 *         value={draft.description}
 *         onChange={e => handleChange('description', e.target.value)}
 *         placeholder="Description"
 *       />
 *       <button type="submit">Submit</button>
 *       <p className="hint">Draft auto-saved for this session</p>
 *     </form>
 *   );
 * }
 * ```
 *
 *
 * @example Resume and complete a session wizard
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';
 *
 * interface WizardState {
 *   currentStep: number;
 *   completedSteps: number[];
 *   data: Record<string, unknown>;
 * }
 *
 * function MultiStepWizard() {
 *   const { value: wizard, setValue: setWizard } = useSPFxSessionStorage<WizardState>('wizard', {
 *     currentStep: 0,
 *     completedSteps: [],
 *     data: {}
 *   });
 *
 *   const goToStep = (step: number) => {
 *     setWizard(prev => ({ ...prev, currentStep: step }));
 *   };
 *
 *   const completeStep = (stepData: Record<string, unknown>) => {
 *     setWizard(prev => ({
 *       currentStep: prev.currentStep + 1,
 *       completedSteps: [...prev.completedSteps, prev.currentStep],
 *       data: { ...prev.data, ...stepData }
 *     }));
 *   };
 *
 *   return (
 *     <div className="wizard">
 *       <nav>
 *         {[0, 1, 2, 3].map(step => (
 *           <button
 *             key={step}
 *             onClick={() => goToStep(step)}
 *             disabled={!wizard.completedSteps.includes(step - 1) && step !== 0}
 *             className={wizard.currentStep === step ? 'active' : ''}
 *           >
 *             Step {step + 1}
 *           </button>
 *         ))}
 *       </nav>
 *       <main>
 *         <p>Current step: {wizard.currentStep + 1}</p>
 *         <button onClick={() => completeStep({ confirmed: true })}>Complete step</button>
 *       </main>
 *     </div>
 *   );
 * }
 * ```
 *
 * @scenario Save and clear an instance scoped session value
 * ```tsx
 * // Requires the matching SPFx provider. Only named commands write; browser JSON persistence is best effort. The string generic does not validate stored JSON. Reset keeps current hook state, rather than rereading persistence; reentry initializes from the scoped key.
 * import * as React from 'react';
 * import { Button, Input } from '@fluentui/react-components';
 * import { useSPFxContext, useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   children?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const { spfxContext } = useSPFxContext();
 *   const { value, setValue, remove } = useSPFxSessionStorage('demo-session-key', '');
 *   const [draft, setDraft] = React.useState('');
 *   const [message, setMessage] = React.useState<string>();
 *   const lifetime = React.useMemo(() => ({ active: true, timer: undefined as number | undefined }),
 *     [resetRevision, spfxContext, setValue, remove]);
 *   React.useLayoutEffect(() => {
 *     setDraft('');
 *     setMessage(undefined);
 *     return () => {
 *       lifetime.active = false;
 *       if (lifetime.timer !== undefined) window.clearTimeout(lifetime.timer);
 *     };
 *   }, [lifetime]);
 *   const save = (): void => {
 *     if (!lifetime.active || !draft) return;
 *     setValue(draft);
 *     setDraft('');
 *     setMessage('Session value updated; persistence attempted.');
 *     if (lifetime.timer !== undefined) window.clearTimeout(lifetime.timer);
 *     const timer = window.setTimeout(() => {
 *       if (!lifetime.active || lifetime.timer !== timer) return;
 *       lifetime.timer = undefined;
 *       setMessage(undefined);
 *     }, 3000);
 *     lifetime.timer = timer;
 *   };
 *   const clear = (): void => { if (lifetime.active) remove(); };
 *   return render({
 *     controls: <Input aria-label="Session value" value={draft} onChange={(_, data) => setDraft(data.value)} />,
 *     commands: <><Button disabled={!draft} onClick={save}>Save Session</Button><Button onClick={clear}>Clear Session</Button></>,
 *     status: message ? <p role="status">{message}</p> : undefined,
 *     children: <output>Current value: {String(value ?? '')}</output>,
 *   });
 * }
 * export const demo = { host: 'webPart', kind: 'write', contentPlacement: 'configuration', observe: 'Save Session updates the current instance value and attempts browser persistence; Clear Session removes its scoped key and restores the empty fallback without new feedback. Reset clears only input/feedback and keeps the current hook value, without rereading browser storage.' };
 * ```
 */
export function useSPFxSessionStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T> {
  return useSPFxBrowserStorage('session', key, defaultValue);
}
