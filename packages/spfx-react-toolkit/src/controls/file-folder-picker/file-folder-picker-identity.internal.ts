import type { FileFolderPickerValue } from './file-folder-picker.types';
import { copyFileFolderPickerValue } from './file-folder-picker-values.internal';

/** Canonical identity excludes kind/display/enrichment snapshots and preserves opaque drive/item ID case. */
export function fileFolderPickerIdentity(value: FileFolderPickerValue): string {
  const verified = copyFileFolderPickerValue(value);
  // Tuple encoding prevents collisions when opaque IDs contain separators.
  return JSON.stringify([verified.tenantId.toLowerCase(), verified.endpoint, verified.driveId, verified.itemId]);
}
