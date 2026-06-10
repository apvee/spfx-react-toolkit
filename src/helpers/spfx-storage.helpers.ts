/**
 * Creates a storage key scoped to an SPFx instance ID.
 *
 * @param instanceId - SPFx component instance ID
 * @param key - Unscoped storage key
 * @returns Storage key scoped to the instance
 */
export function createScopedSPFxStorageKey(instanceId: string, key: string): string {
  return 'spfx:' + instanceId + ':' + key;
}
