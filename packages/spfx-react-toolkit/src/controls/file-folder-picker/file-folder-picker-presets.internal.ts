import type { FileFolderPickerPreset, FileFolderPickerProps, FileFolderPickerResolvedOptions, FileFolderPickerScope } from './file-folder-picker.types';
import { normalizeFileFolderPickerItemPolicy } from './file-folder-picker-results.internal';
import { normalizeFileFolderPickerEndpoint } from './file-folder-picker-values.internal';
import { validateFileFolderPickerFolderPath } from './file-folder-picker-scope.internal';
import { pickerListResource } from '../pickers/picker-list-target.internal';

/** Opening-owned configuration shared by SDK compilation and authentication/session policy. */
export interface ResolvedFileFolderPickerExperience {
  readonly mode: 'legacy' | 'preset';
  readonly preset?: FileFolderPickerPreset;
  /** Copy of common configuration and effective SDK settings; no new selection or remote metadata. */
  readonly options: FileFolderPickerResolvedOptions;
  readonly authentication: {
    readonly graphEnabled: boolean;
    readonly claimsChallengeEnabled: boolean;
    /** Only permits lazy discovery; the session must independently verify the current user's resource. */
    readonly auxiliaryPersonalResourceEnabled: boolean;
  };
}

// Finite keys are deliberately checked by defined value, including false and empty arrays.
const legacyKeys = [
  'graphEnabled', 'claimsChallengeEnabled', 'navigationPreset', 'visibleLocations', 'initialView',
  'navigationVisible', 'navigationInitialState', 'siteNavigationUrl', 'fromLinkMode', 'localFilesMode',
  'localFilesLabel', 'stickyHeaders', 'retainSelectionOnNavigate', 'columnFilterMode', 'trayPrompt',
  'initialLocationFallback', 'requiredAccess', 'title', 'selectLabel', 'cancelLabel', 'layout', 'initialSort', 'initialFilter',
] as const;

function choice<T extends string>(value: T | undefined, allowed: readonly T[]): void {
  if (value !== undefined && !allowed.includes(value)) throw new Error('Invalid picker configuration option.');
}
function label(value: string | undefined): void {
  if (value !== undefined && (typeof value !== 'string' || !value.trim())) throw new Error('Empty native label.');
}
function sourceUrl(value: string): string {
  try {
    if (typeof value !== 'string') throw new Error();
    const endpoint = normalizeFileFolderPickerEndpoint(`${value.replace(/\/$/, '')}/_api/v2.0`);
    return endpoint.slice(0, -'/_api/v2.0'.length);
  } catch { throw new Error('Invalid picker source URL.'); }
}
function validateScope(scope: FileFolderPickerScope | undefined, props: FileFolderPickerProps): void {
  if (scope !== undefined && (typeof scope !== 'object' || scope === null || Array.isArray(scope) || scope.kind === undefined)) throw new Error('Invalid picker scope.');
  choice(scope?.kind, ['siteCollection', 'library', 'folder', 'oneDrive']);
  const oneDrive = scope?.kind === 'oneDrive';
  let webUrl = 'https://picker.sharepoint.com';
  if (scope?.kind === 'siteCollection' && scope.siteUrl !== undefined) webUrl = sourceUrl(scope.siteUrl);
  if (scope?.kind === 'library' || scope?.kind === 'folder') {
    webUrl = sourceUrl(scope.webUrl);
    try { pickerListResource(scope.list, webUrl); } catch { throw new Error('Invalid picker library selector.'); }
  }
  if (scope?.kind === 'folder') validateFileFolderPickerFolderPath(scope.folder);
  if (scope?.kind === 'oneDrive' && scope.folder !== undefined) validateFileFolderPickerFolderPath(scope.folder);
  const initial = props.initialLocation;
  if (initial !== undefined) {
    if (typeof initial !== 'object' || initial === null || Array.isArray(initial)) throw new Error('Invalid initial location.');
    if (oneDrive && (initial.webUrl !== undefined || initial.list !== undefined)) throw new Error('OneDrive initial location only accepts a folder.');
    if (!oneDrive && initial.folder !== undefined && initial.list === undefined) throw new Error('Initial SharePoint folder requires a list.');
    if (initial.webUrl !== undefined) webUrl = sourceUrl(initial.webUrl);
    if (initial.list !== undefined) {
      try { pickerListResource(initial.list, webUrl); } catch { throw new Error('Invalid picker initial library selector.'); }
    }
    if (initial.folder !== undefined) validateFileFolderPickerFolderPath(initial.folder);
  }
  // Discovery entry must not replace a constrained folder with an unrelated initial view.
  if (props.initialView !== undefined && props.initialView !== 'location'
    && (!oneDrive || props.initialLocation !== undefined || scope.folder !== undefined)) throw new Error('Initial view conflicts with source or location.');
}
function validateOptions(props: FileFolderPickerProps): void {
  for (const value of [props.multiselect, props.searchable, props.stickyHeaders, props.navigationVisible,
    props.retainSelectionOnNavigate, props.graphEnabled, props.claimsChallengeEnabled, props.uploadEnabled,
    props.createFolderEnabled, props.initialSort?.ascending]) {
    if (value !== undefined && typeof value !== 'boolean') throw new Error('Invalid native boolean option.');
  }
  if (props.extensions !== undefined && !Array.isArray(props.extensions)) throw new Error('Invalid picker file extensions.');
  if (props.itemTypes !== undefined && !['files', 'folders', 'all'].includes(props.itemTypes)) throw new Error('Invalid picker item mode.');
  normalizeFileFolderPickerItemPolicy(props);
  for (const value of [props.title, props.selectLabel, props.cancelLabel, props.localFilesLabel]) label(value);
  choice(props.layout, ['details', 'compact-details', 'tiles']);
  choice(props.requiredAccess, ['read', 'read-write']);
  choice(props.navigationInitialState, ['modal', 'hidden']);
  choice(props.navigationPreset, ['oneDrive', 'current-site']);
  choice(props.columnFilterMode, ['panel', 'menu']);
  choice(props.trayPrompt, ['selection-summary', 'selection-editor', 'none']);
  choice(props.fromLinkMode, ['disabled', 'navigation', 'pivot']);
  choice(props.localFilesMode, ['disabled', 'upload']);
  choice(props.initialLocationFallback, ['error', 'scopeRoot']);
  choice(props.initialView, ['location', 'myFiles', 'recent', 'shared', 'site', 'organization']);
  if (props.visibleLocations !== undefined && !Array.isArray(props.visibleLocations)) throw new Error('Invalid picker navigation locations.');
  for (const location of props.visibleLocations ?? []) {
    if (typeof location !== 'string') throw new Error('Invalid picker navigation location.');
    choice(location, ['myFiles', 'recent', 'shared', 'quickAccess', 'site', 'organization']);
  }
  if (props.initialSort !== undefined && (props.initialSort === null || typeof props.initialSort !== 'object' || typeof props.initialSort.fieldName !== 'string' || !props.initialSort.fieldName.trim())) throw new Error('Initial sort requires fieldName.');
  if (props.initialFilter !== undefined && (props.initialFilter === null || typeof props.initialFilter !== 'object' || typeof props.initialFilter.fieldName !== 'string' || !props.initialFilter.fieldName.trim() || typeof props.initialFilter.value !== 'string')) throw new Error('Initial filter requires fieldName and value.');
  if (props.localFilesMode === 'upload' && props.uploadEnabled !== true) throw new Error('Device upload requires uploadEnabled.');
  if (props.siteNavigationUrl !== undefined) sourceUrl(props.siteNavigationUrl);
  validateScope(props.scope, props);
}
function copyOptions(props: FileFolderPickerProps): FileFolderPickerResolvedOptions {
  // Models retain their private verification stamps; only configuration is copied.
  return { ...props,
    ...(props.scope === undefined ? {} : { scope: {
      ...props.scope,
      ...(('list' in props.scope && typeof props.scope.list === 'object') ? { list: { ...props.scope.list } } : {}),
    } }),
    ...(props.initialLocation === undefined ? {} : { initialLocation: {
      ...props.initialLocation,
      ...(typeof props.initialLocation.list === 'object' ? { list: { ...props.initialLocation.list } } : {}),
    } }),
    ...(props.extensions === undefined ? {} : { extensions: [...props.extensions] }),
    ...(props.visibleLocations === undefined ? {} : { visibleLocations: [...props.visibleLocations] }),
    ...(props.initialSort === undefined ? {} : { initialSort: { ...props.initialSort } }),
    ...(props.initialFilter === undefined ? {} : { initialFilter: { ...props.initialFilter } }),
  };
}
/** Pure validation and recipe resolution. No tokens, network reads, persistence or provider creation.
 * Undefined preset preserves legacy values/defaults exactly. Defined legacy options cannot be silently ignored.
 * Containment and source identity that require authoritative metadata remain the scope/results resolvers' job.
 */
export function resolveFileFolderPickerExperience(props: FileFolderPickerProps): ResolvedFileFolderPickerExperience {
  choice(props.preset, ['focused', 'extended']);
  if (props.preset !== undefined && legacyKeys.some(key => props[key] !== undefined)) throw new Error('A picker preset cannot be combined with legacy options.');
  validateOptions(props);
  const options = copyOptions(props);
  if (props.preset === undefined) return {
    mode: 'legacy', options,
    authentication: {
      graphEnabled: props.graphEnabled ?? false,
      claimsChallengeEnabled: props.claimsChallengeEnabled ?? false,
      auxiliaryPersonalResourceEnabled: props.scope?.kind !== 'oneDrive' && props.navigationVisible !== false
        && props.visibleLocations?.some(location => location === 'myFiles' || location === 'recent' || location === 'shared') === true,
    },
  };
  const policy = normalizeFileFolderPickerItemPolicy(props);
  const extended = props.preset === 'extended', oneDrive = props.scope?.kind === 'oneDrive';
  return {
    mode: 'preset', preset: props.preset,
    options: {
      ...options, itemTypes: policy.itemTypes, multiselect: policy.multiselect, extensions: policy.extensions,
      searchable: props.searchable ?? true, uploadEnabled: props.uploadEnabled ?? false, createFolderEnabled: props.createFolderEnabled ?? false,
      initialView: 'location', initialLocationFallback: 'error', navigationVisible: extended,
      navigationPreset: oneDrive ? 'oneDrive' : 'current-site',
      visibleLocations: extended ? (oneDrive ? ['myFiles', 'recent', 'shared', 'organization', 'site'] : ['recent', 'shared', 'quickAccess', 'organization', 'site']) : [],
      fromLinkMode: extended ? 'pivot' : 'disabled', localFilesMode: props.uploadEnabled === true ? 'upload' : 'disabled',
      graphEnabled: true, claimsChallengeEnabled: true, requiredAccess: 'read', trayPrompt: 'selection-summary',
      retainSelectionOnNavigate: policy.multiselect,
    },
    authentication: { graphEnabled: true, claimsChallengeEnabled: true, auxiliaryPersonalResourceEnabled: extended && !oneDrive },
  };
}
