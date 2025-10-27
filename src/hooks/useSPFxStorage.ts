// useSPFxStorage.ts
// Hooks for persisted storage scoped to SPFx instance

import { atomWithStorage } from 'jotai/utils';
import { useAtom } from 'jotai';
import { useMemo } from 'react';
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

/**
 * Hook to use localStorage scoped to SPFx instance
 * 
 * Creates a persisted state atom using Jotai's atomWithStorage.
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
  const { id: instanceId } = useSPFxInstanceInfo();
  
  // Create scoped storage key
  const scopedKey = 'spfx:' + instanceId + ':' + key;
  
  // Create atom with storage (memoized to avoid recreation)
  const storageAtom = useMemo(
    () => atomWithStorage<T>(scopedKey, defaultValue),
    [scopedKey, defaultValue]
  );
  
  const [value, setValue] = useAtom(storageAtom);
  
  // Remove function (reset to default)
  const remove = (): void => {
    setValue(defaultValue);
    if (typeof localStorage !== 'undefined') {
      localStorage.removeItem(scopedKey);
    }
  };
  
  return {
    value,
    setValue,
    remove,
  };
}

/**
 * Hook to use sessionStorage scoped to SPFx instance
 * 
 * Creates a persisted state atom using Jotai's atomWithStorage.
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
  const { id: instanceId } = useSPFxInstanceInfo();
  
  // Create scoped storage key
  const scopedKey = 'spfx:' + instanceId + ':' + key;
  
  // Create atom with session storage
  const storageAtom = useMemo(
    () => atomWithStorage<T>(
      scopedKey,
      defaultValue,
      {
        getItem: (key) => {
          if (typeof sessionStorage === 'undefined') {
            return defaultValue;
          }
          const item = sessionStorage.getItem(key);
          if (item === null) {
            return defaultValue;
          }
          try {
            return JSON.parse(item) as T;
          } catch {
            return defaultValue;
          }
        },
        setItem: (key, value) => {
          if (typeof sessionStorage !== 'undefined') {
            sessionStorage.setItem(key, JSON.stringify(value));
          }
        },
        removeItem: (key) => {
          if (typeof sessionStorage !== 'undefined') {
            sessionStorage.removeItem(key);
          }
        },
      }
    ),
    [scopedKey, defaultValue]
  );
  
  const [value, setValue] = useAtom(storageAtom);
  
  // Remove function (reset to default)
  const remove = (): void => {
    setValue(defaultValue);
    if (typeof sessionStorage !== 'undefined') {
      sessionStorage.removeItem(scopedKey);
    }
  };
  
  return {
    value,
    setValue,
    remove,
  };
}
