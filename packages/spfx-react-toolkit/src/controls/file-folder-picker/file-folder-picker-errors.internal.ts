import { resolveFileFolderPickerStrings } from './file-folder-picker-strings.internal';
import type { FileFolderPickerStrings } from './file-folder-picker-strings.types';
/** Stable, sanitized categories shared by the protocol, session and outer Field. */
export type FileFolderPickerErrorCode = 'auth-resource' | 'auth-type' | 'auth-challenge' | 'auth-provider'
  | 'selection-identity' | 'selection-scope' | 'selection-policy' | 'selection-unavailable' | 'selection-invalid'
  | 'configuration' | 'opening' | 'onedrive-unavailable' | 'timeout' | 'communication';
export class FileFolderPickerError extends Error {
  constructor(readonly code: FileFolderPickerErrorCode, strings: FileFolderPickerStrings = resolveFileFolderPickerStrings()) {
    super(strings.errors[code]);
    this.name = 'FileFolderPickerError';
  }
}
/** Do not expose provider/native/HTTP payload strings through Field or onError. */
export function fileFolderPickerFailure(failure: unknown, fallback: FileFolderPickerErrorCode, strings: FileFolderPickerStrings = resolveFileFolderPickerStrings()): FileFolderPickerError {
  if (failure instanceof FileFolderPickerError) return new FileFolderPickerError(failure.code, strings);
  if (fallback !== 'selection-invalid') return new FileFolderPickerError(fallback, strings);
  const detail = failure instanceof Error ? failure.message : '';
  if (/identity mismatch|drive metadata mismatch|provenance|ambiguous.*metadata/i.test(detail)) return new FileFolderPickerError('selection-identity', strings);
  if (/outside|another tenant|crossed a drive|untrusted.*endpoint/i.test(detail)) return new FileFolderPickerError('selection-scope', strings);
  if (/not allowed|selection exceeds|limit|extension|category/i.test(detail)) return new FileFolderPickerError('selection-policy', strings);
  if (/deleted|unavailable|\b403\b|\b404\b/i.test(detail)) return new FileFolderPickerError('selection-unavailable', strings);
  return new FileFolderPickerError(fallback, strings);
}
