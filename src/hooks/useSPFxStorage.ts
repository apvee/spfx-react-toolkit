// useSPFxStorage.ts
// Hooks for persisted storage scoped to SPFx instance

import { useMemo, useCallback, useEffect, useState } from 'react';
import { createScopedSPFxStorageKey } from '../helpers/spfx-storage.helpers';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';

/**
 * Return type for storage hooks
 */
export interface SPFxStorageHook<T> {
  /** Current value */
  readonly value: T;
  
  /** Set new value */
  readonly setValue: (value: T | ((prev: T) => T)) => void;
  
  /** Remove value (reset to default) */
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

  const [value, setStoredValue] = useState<T>(() => (
    readStorageValue(getBrowserStorage(kind), scopedKey, defaultValue)
  ));

  useEffect(() => {
    setStoredValue(readStorageValue(getBrowserStorage(kind), scopedKey, defaultValue));
  }, [kind, scopedKey, defaultValue]);

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
    setStoredValue(defaultValue);
  }, [kind, scopedKey, defaultValue]);

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

      setStoredValue(readStorageValue(storage, scopedKey, defaultValue));
    };

    window.addEventListener('storage', handleStorage);

    return () => {
      window.removeEventListener('storage', handleStorage);
    };
  }, [kind, scopedKey, defaultValue]);

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
 * @param key - Storage key (will be prefixed with instance ID)
 * @param defaultValue - Default value if not in storage
 * @returns Storage hook with value, setValue, and remove
 * 
 * @example
 * ```tsx
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
 * @param key - Storage key (will be prefixed with instance ID)
 * @param defaultValue - Default value if not in storage
 * @returns Storage hook with value, setValue, and remove
 * 
 * @example
 * ```tsx
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
 */
export function useSPFxSessionStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T> {
  return useSPFxBrowserStorage('session', key, defaultValue);
}
