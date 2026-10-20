import type { FileFolderPickerInitialLocation, FileFolderPickerPreset, FileFolderPickerProps, FileFolderPickerScope } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';

/** Saved SDK-facing fields are read only for compatibility, never exposed as configurator switches. */
const savedNavigation = [
  ['myFiles', 'fileFolderMyFiles'], ['recent', 'fileFolderRecent'], ['shared', 'fileFolderShared'],
  ['quickAccess', 'fileFolderQuickAccess'], ['organization', 'fileFolderOrganization'], ['site', 'fileFolderSite'],
] as const;

type LegacyProps = Extract<FileFolderPickerProps, { readonly preset?: undefined }>;
export interface FileFolderDemoSettings extends Partial<Record<typeof savedNavigation[number][1], boolean>> {
  fileFolderPreset?: FileFolderPickerPreset | 'legacy';
  fileFolderTitle?: string;
  fileFolderLayout?: LegacyProps['layout'];
  fileFolderGraph?: boolean;
  fileFolderNavigationVisible?: boolean;
  fileFolderInitialView?: LegacyProps['initialView'];
  fileFolderNavigationPreset?: LegacyProps['navigationPreset'] | 'auto';
  fileFolderSiteUrl?: string;
  fileFolderFromLink?: LegacyProps['fromLinkMode'];
  fileFolderDeviceUpload?: boolean;
  fileFolderDeviceLabel?: string;
  fileFolderUpload?: boolean;
  fileFolderCreateFolder?: boolean;
}
const savedSdkProperties = [
  ...savedNavigation.map(([, property]) => property), 'fileFolderTitle', 'fileFolderLayout', 'fileFolderGraph',
  'fileFolderNavigationVisible', 'fileFolderInitialView', 'fileFolderNavigationPreset', 'fileFolderSiteUrl',
  'fileFolderFromLink', 'fileFolderDeviceUpload', 'fileFolderDeviceLabel',
] as const;

/** Defined false and empty saved settings are meaningful compatibility preferences. Writes remain common intents. */
export function fileFolderHasLegacySettings(settings: FileFolderDemoSettings): boolean {
  return savedSdkProperties.some(property => settings[property] !== undefined);
}
export function fileFolderDemoPreset(settings: FileFolderDemoSettings): FileFolderPickerPreset | 'legacy' {
  return settings.fileFolderPreset ?? (fileFolderHasLegacySettings(settings) ? 'legacy' : 'focused');
}

type LegacyExperience = Pick<LegacyProps, 'title' | 'layout' | 'graphEnabled' | 'visibleLocations' | 'initialView' | 'navigationVisible' | 'navigationPreset' | 'siteNavigationUrl' | 'fromLinkMode' | 'localFilesMode' | 'localFilesLabel'> & { readonly preset?: undefined };
/** Select one library contract; recipe generation stays in the library resolver. Original saved fields remain restorable. */
export function fileFolderDemoExperience(settings: FileFolderDemoSettings, scope: FileFolderPickerScope, initialLocation?: FileFolderPickerInitialLocation): LegacyExperience | { readonly preset: FileFolderPickerPreset } {
  const preset = fileFolderDemoPreset(settings);
  if (preset !== 'legacy') return { preset };
  const visibleLocations = savedNavigation.every(([, property]) => settings[property] === undefined) ? undefined
    : savedNavigation.filter(([key, property]) => {
      const incompatibleSource = (key === 'myFiles' && scope.kind !== 'oneDrive') || (key === 'quickAccess' && scope.kind === 'oneDrive');
      return !incompatibleSource && (settings[property] ?? (key === 'myFiles' && scope.kind === 'oneDrive'));
    }).map(([key]) => key);
  return {
    title: settings.fileFolderTitle || undefined, layout: settings.fileFolderLayout, graphEnabled: settings.fileFolderGraph,
    visibleLocations, navigationVisible: settings.fileFolderNavigationVisible,
    initialView: scope.kind === 'oneDrive' && !scope.folder && initialLocation === undefined ? settings.fileFolderInitialView ?? 'location' : 'location',
    navigationPreset: settings.fileFolderNavigationPreset === 'oneDrive' || settings.fileFolderNavigationPreset === 'current-site' ? settings.fileFolderNavigationPreset : undefined,
    siteNavigationUrl: settings.fileFolderSiteUrl?.trim() || undefined, fromLinkMode: settings.fileFolderFromLink,
    localFilesMode: settings.fileFolderUpload && settings.fileFolderDeviceUpload ? 'upload' : 'disabled', localFilesLabel: settings.fileFolderDeviceLabel || undefined,
  };
}
