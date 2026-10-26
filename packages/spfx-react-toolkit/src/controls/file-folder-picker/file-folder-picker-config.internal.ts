import type { FileFolderPickerOption, FileFolderPickerProps } from './file-folder-picker.types';
import type { ResolvedFileFolderPickerScope } from './file-folder-picker-scope.internal';
import { fileFolderPickerOptionMatchesPolicy, isVerifiedFileFolderPickerOption, normalizeFileFolderPickerItemPolicy } from './file-folder-picker-results.internal';

import { resolveFileFolderPickerExperience } from './file-folder-picker-presets.internal';
import type { ResolvedFileFolderPickerExperience } from './file-folder-picker-presets.internal';

type DistributiveOmit<T, Key extends PropertyKey> = T extends unknown ? Omit<T, Key> : never;

interface FilePickerPath { readonly web: string; readonly list?: string; readonly folder?: string; }
export interface FilePickerV8SourceItem { readonly id: string; readonly parentReference: { readonly driveId: string }; readonly '@sharePoint.endpoint': string; }
/** Flat consumer options plus resolver-owned preselection; display seeds are never admitted. */
export type FileFolderPickerConfigurationProps = DistributiveOmit<FileFolderPickerProps, 'value' | 'defaultValue' | 'onChange'> & {
  readonly sourceItems?: readonly FileFolderPickerOption[];
};
export function compileFileFolderPickerConfig(scope: ResolvedFileFolderPickerScope, input: FileFolderPickerConfigurationProps, channel: { readonly origin: string; readonly channelId: string; readonly theme?: 'default' | 'dark' }, experience: ResolvedFileFolderPickerExperience = resolveFileFolderPickerExperience(input)): FilePickerV8Configuration {
  const props = experience.options;
  const policy = normalizeFileFolderPickerItemPolicy(props);
  const filters: NonNullable<FilePickerV8Configuration['typesAndSources']['filters']> = policy.fileCategory === undefined ? policy.extensions : [policy.fileCategory === 'images' ? 'photo' : policy.fileCategory === 'documents' ? 'document' : policy.fileCategory, ...(policy.itemTypes === 'all' ? ['folder' as const] : [])];
  const origin = new URL(channel.origin);
  if (!['http:', 'https:'].includes(origin.protocol) || origin.origin !== channel.origin || !channel.channelId.trim()) throw new Error('Invalid picker messaging channel.');
  const initial: FilePickerPath = { web: scope.initial.webUrl, ...(scope.initial.listPath === undefined ? {} : { list: scope.initial.listPath }), ...(scope.initial.folderPath === undefined ? {} : { folder: scope.initial.folderPath }) };
  const restriction: FilePickerPath = { web: scope.webUrl, ...(scope.library === undefined ? {} : { list: scope.library.rootFolderPath }), ...(scope.folder === undefined ? {} : { folder: scope.folder.serverRelativePath }) };
  const seen = new Set<string>();
  const sourceItems = (input.sourceItems ?? []).filter(option => isVerifiedFileFolderPickerOption(option, scope) && fileFolderPickerOptionMatchesPolicy(option, props)).filter(option => {
    const key = JSON.stringify([option.value.endpoint, option.value.driveId, option.value.itemId]);
    if (seen.has(key)) return false; seen.add(key); return true;
  }).map(option => ({ id: option.value.itemId, parentReference: { driveId: option.value.driveId }, '@sharePoint.endpoint': option.value.endpoint }));
  // Incompatible outer values remain visible to the caller. Omit all native preselection rather than truncate them.
  const preselection = policy.maximumCount !== undefined && sourceItems.length > policy.maximumCount ? [] : sourceItems;
  const oneDrive = scope.kind === 'oneDrive';
  const { title, selectLabel, cancelLabel } = props;
  const visible = props.visibleLocations ?? (oneDrive ? ['myFiles'] : []);
  const siteUrl = props.siteNavigationUrl ?? scope.webUrl;
  const site = new URL(siteUrl), collection = new URL(scope.siteCollectionUrl);
  const knownSites = [scope.webUrl, scope.siteCollectionUrl, scope.initial.webUrl].map(url => new URL(url).href.replace(/\/$/, '').toLowerCase());
  if (site.origin !== collection.origin || site.username || site.password || site.search || site.hash || !knownSites.includes(site.href.replace(/\/$/, '').toLowerCase())) throw new Error('Site navigation requires a resolved web in the allowed collection.');
  const view = props.initialView ?? 'location';
  const entry = view === 'location' ? (oneDrive
    ? { oneDrive: { files: { ...(scope.initial.folderPath === undefined ? {} : { folder: scope.initial.folderPath }), fallbackToRoot: false } } }
    : { sharePoint: { byPath: { ...initial, fallbackToRoot: false } } })
    : view === 'site' ? { site: {} } : view === 'organization' ? { myOrganization: {} }
    : { oneDrive: view === 'recent' ? { recent: {} } : view === 'shared' ? { sharedWithMe: {} } : { files: { fallbackToRoot: false } } };
  return {
    sdk: '8.0' as const, messaging: { origin: channel.origin, channelId: channel.channelId },
    entry: { ...entry,
      ...(props.initialSort === undefined ? {} : { sortBy: { fieldName: props.initialSort.fieldName, isAscending: props.initialSort.ascending ?? true } }),
      ...(props.initialFilter === undefined ? {} : { filterBy: { ...props.initialFilter } }),
    },
    ...(title === undefined ? {} : { title }),
    theme: channel.theme ?? 'default',
    list: { ...(props.layout === undefined ? {} : { layout: { type: props.layout } }), ...(props.stickyHeaders === undefined ? {} : { scrolling: { enableStickyHeaders: props.stickyHeaders } }) },
    leftNav: { enabled: props.navigationVisible ?? true, preset: props.navigationPreset ?? (oneDrive ? 'oneDrive' : 'current-site'), ...(props.navigationInitialState === undefined ? {} : { initialModality: props.navigationInitialState }) },
    tray: { prompt: props.trayPrompt ?? 'selection-summary' },
    ...(visible.includes('site') || view === 'site' ? { pivots: { site: { byPath: { web: siteUrl } } } } : {}),
    search: { enabled: props.searchable ?? true },
    authentication: { enabled: true, tokens: { sharePoint: true, graph: experience.authentication.graphEnabled, substrate: false }, claimsChallenge: { enabled: experience.authentication.claimsChallengeEnabled } },
    typesAndSources: { mode: policy.itemTypes, ...(filters.length === 0 ? {} : { filters }), locations: oneDrive ? { oneDrive: {} } : { sharePoint: { byPath: restriction } }, access: { mode: props.requiredAccess ?? 'read' }, pivots: { oneDrive: visible.includes('myFiles'), recent: visible.includes('recent'), shared: visible.includes('shared'), sharedLibraries: visible.includes('quickAccess'), myOrganization: visible.includes('organization'), site: visible.includes('site') } },
    selection: { mode: policy.multiselect ? 'multiple' : 'single', ...(policy.maximumCount === undefined ? {} : { maximumCount: policy.maximumCount }), ...(props.retainSelectionOnNavigate === undefined ? (policy.multiselect ? { enablePersistence: true } : {}) : { enablePersistence: props.retainSelectionOnNavigate }), ...(preselection.length === 0 ? {} : { sourceItems: preselection }) },
    commands: {
      pick: { action: 'select' as const, ...(selectLabel === undefined ? {} : { label: selectLabel }), select: { urls: { download: false } } },
      ...(cancelLabel === undefined ? {} : { close: { label: cancelLabel } }),
      upload: { enabled: props.uploadEnabled ?? false }, createFolder: { enabled: props.createFolderEnabled ?? false },
      browseThisDevice: props.localFilesMode === 'upload' ? { enabled: true, mode: 'upload' as const, ...(props.localFilesLabel === undefined ? {} : { label: props.localFilesLabel }) } : { enabled: false },
      fromALink: props.fromLinkMode === undefined || props.fromLinkMode === 'disabled' ? { enabled: false } : { enabled: true, mode: props.fromLinkMode === 'navigation' ? 'nav' as const : 'pivot' as const },
      switchAccount: { mode: 'none' as const }, manageAccounts: { mode: 'none' as const },
      ...(props.columnFilterMode === undefined ? {} : { filterByColumn: { mode: props.columnFilterMode } }),
    },
  };
}
/** Tested SPFx-compatible v8 configuration. Unknown protocol/result families are not forwarded. */
export interface FilePickerV8Configuration {
  readonly sdk: '8.0';
  readonly messaging: { readonly origin: string; readonly channelId: string };
  readonly entry: {
    readonly sharePoint?: { readonly byPath: FilePickerPath & { readonly fallbackToRoot: boolean } };
    readonly oneDrive?: { readonly files?: { readonly folder?: string; readonly fallbackToRoot: boolean }; readonly recent?: Readonly<Record<string, never>>; readonly sharedWithMe?: Readonly<Record<string, never>> };
    readonly site?: Readonly<Record<string, never>>;
    readonly myOrganization?: Readonly<Record<string, never>>;
    readonly sortBy?: { readonly fieldName: string; readonly isAscending: boolean };
    readonly filterBy?: { readonly fieldName: string; readonly value: string };
  };
  readonly title?: string;
  readonly theme: 'default' | 'dark';
  readonly list: { readonly layout?: { readonly type: 'details' | 'compact-details' | 'tiles' }; readonly scrolling?: { readonly enableStickyHeaders: boolean } };
  readonly leftNav: { readonly enabled: boolean; readonly preset: 'oneDrive' | 'current-site'; readonly initialModality?: 'modal' | 'hidden' };
  readonly tray: { readonly prompt: 'selection-summary' | 'selection-editor' | 'none' };
  readonly pivots?: { readonly site: { readonly byPath: { readonly web: string } } };
  readonly search: { readonly enabled: boolean };
  readonly authentication: { readonly enabled: boolean; readonly tokens: { readonly sharePoint: boolean; readonly graph: boolean; readonly substrate: boolean }; readonly claimsChallenge: { readonly enabled: boolean } };
  readonly typesAndSources: {
    readonly mode: 'files' | 'folders' | 'all'; readonly filters?: readonly (`.${string}` | 'photo' | 'video' | 'audio' | 'media' | 'document' | 'folder')[];
    readonly locations: { readonly sharePoint?: { readonly byPath: FilePickerPath }; readonly oneDrive?: Readonly<Record<string, never>> };
    readonly access: { readonly mode: 'read' | 'read-write' };
    readonly pivots: { readonly oneDrive: boolean; readonly recent: boolean; readonly shared: boolean; readonly sharedLibraries: boolean; readonly myOrganization: boolean; readonly site: boolean };
  };
  readonly selection: { readonly mode: 'single' | 'multiple'; readonly maximumCount?: number; readonly enablePersistence?: boolean; readonly sourceItems?: readonly FilePickerV8SourceItem[] };
  readonly commands: {
    readonly pick: { readonly action: 'select'; readonly label?: string; readonly select: { readonly urls: { readonly download: false } } };
    readonly close?: { readonly label: string };
    readonly upload: { readonly enabled: boolean }; readonly createFolder: { readonly enabled: boolean };
    readonly browseThisDevice: { readonly enabled: boolean; readonly mode?: 'upload'; readonly label?: string };
    readonly fromALink: { readonly enabled: boolean; readonly mode?: 'nav' | 'pivot' };
    readonly switchAccount: { readonly mode: 'none' }; readonly manageAccounts: { readonly mode: 'none' };
    readonly filterByColumn?: { readonly mode: 'panel' | 'menu' };
  };
}
