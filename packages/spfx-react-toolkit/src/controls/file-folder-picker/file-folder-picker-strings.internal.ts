import type { FileFolderPickerStrings, FileFolderPickerStringsOverride } from './file-folder-picker-strings.types';
import { nativeFileFolderPickerResource } from './localization/file-folder-picker-native.internal';
import { formatResourceTemplate } from '../../localization/resource-formatters.internal';

const resource = nativeFileFolderPickerResource;
const nativeStrings: FileFolderPickerStrings = {
  browse: {
    files: resource.browse.files,
    folders: resource.browse.folders,
    all: resource.browse.all,
  },
  placeholders: {
    files: resource.placeholders.files,
    folders: resource.placeholders.folders,
    all: resource.placeholders.all,
  },
  labels: {
    selection: resource.labels.selection,
    file: resource.labels.file,
    folder: resource.labels.folder,
    unresolvedItem: resource.labels.unresolvedItem,
    removeItem: (name: string) => formatResourceTemplate(resource.labels.removeItem, { name }),
  },
  validation: {
    invalidReferences: resource.validation.invalidReferences,
    controlledModeChanged: resource.validation.controlledModeChanged,
    incompatibleItems: resource.validation.incompatibleItems,
    invalidPolicy: resource.validation.invalidPolicy,
    selectionLimitExceeded: resource.validation.selectionLimitExceeded,
  },
  errors: {
    configuration: resource.errors.configuration,
    'auth-resource': resource.errors['auth-resource'],
    'auth-type': resource.errors['auth-type'],
    'auth-challenge': resource.errors['auth-challenge'],
    'auth-provider': resource.errors['auth-provider'],
    'selection-identity': resource.errors['selection-identity'],
    'selection-scope': resource.errors['selection-scope'],
    'selection-policy': resource.errors['selection-policy'],
    'selection-unavailable': resource.errors['selection-unavailable'],
    'selection-invalid': resource.errors['selection-invalid'],
    'onedrive-unavailable': resource.errors['onedrive-unavailable'],
    opening: resource.errors.opening,
    timeout: resource.errors.timeout,
    communication: resource.errors.communication,
  },
  protocol: {
    optionalAuthenticationTimeout: resource.protocol.optionalAuthenticationTimeout,
    unsupportedCommand: resource.protocol.unsupportedCommand,
    optionalGraphAuthenticationFailed: resource.protocol.optionalGraphAuthenticationFailed,
    optionalPersonalAuthenticationFailed: resource.protocol.optionalPersonalAuthenticationFailed,
  },
};

/** Known-field merge retains explicit empty values and ignores extra JavaScript keys. */
function mergeFields<T extends object>(defaults: T, overrides?: Partial<T>): T {
  const result = { ...defaults };
  for (const key of Object.keys(defaults) as Array<keyof T>) {
    const value = overrides?.[key];
    if (value !== undefined) result[key] = value;
  }
  return result;
}
export function resolveFileFolderPickerStrings(overrides?: FileFolderPickerStringsOverride): FileFolderPickerStrings {
  return {
    browse: mergeFields(nativeStrings.browse, overrides?.browse),
    placeholders: mergeFields(nativeStrings.placeholders, overrides?.placeholders),
    labels: mergeFields(nativeStrings.labels, overrides?.labels),
    validation: mergeFields(nativeStrings.validation, overrides?.validation),
    errors: mergeFields(nativeStrings.errors, overrides?.errors),
    protocol: mergeFields(nativeStrings.protocol, overrides?.protocol),
  };
}
