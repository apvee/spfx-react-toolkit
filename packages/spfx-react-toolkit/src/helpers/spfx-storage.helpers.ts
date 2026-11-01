/**
 * Storage-name prefixes scoped to a supplied SPFx instance ID.
 *
 * @packageDocumentation
 * @docPage helpers/storage
 */
/**
 * Creates `spfx:<instanceId>:<key>` by concatenating both required strings unchanged. Empty strings, whitespace and colons are retained. This only namespaces a key; it performs no storage I/O, uniqueness check, access control or encryption. Keep both inputs stable for later reads.
 *
 * @param instanceId - SPFx component instance ID
 * @param key - Unscoped storage key
 * @returns Storage key scoped to the instance
 *
 * @example
 * ```ts
 * import { createScopedSPFxStorageKey } from '@apvee/spfx-react-toolkit';
 *
 * const key = createScopedSPFxStorageKey('webpart-42', 'filters'); // spfx:webpart-42:filters
 * ```
 * @useWhen You persist per-instance browser settings and need the same name prefix on subsequent reads.
 */
export function createScopedSPFxStorageKey(instanceId: string, key: string): string {
  return 'spfx:' + instanceId + ':' + key;
}
