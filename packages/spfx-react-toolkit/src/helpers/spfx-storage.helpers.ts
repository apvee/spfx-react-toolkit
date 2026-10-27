/**
 * Creates a storage key scoped to an SPFx instance ID.
 *
 * @param instanceId - SPFx component instance ID
 * @param key - Unscoped storage key
 * @returns Storage key scoped to the instance
 *
 * @example
 * ```ts
 * const key = createScopedSPFxStorageKey('webpart-42', 'filters'); // spfx:webpart-42:filters
 * ```
 */
export function createScopedSPFxStorageKey(instanceId: string, key: string): string {
  return 'spfx:' + instanceId + ':' + key;
}
