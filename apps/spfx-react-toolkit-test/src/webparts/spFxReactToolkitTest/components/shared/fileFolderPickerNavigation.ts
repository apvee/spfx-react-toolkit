import type { FileFolderPickerPreset, FileFolderPickerProps } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';

/** Private saved property schema; runtime compatibility mapping belongs to the source-owned integrated native example. */
type LegacyProps = Extract<FileFolderPickerProps, { readonly preset?: undefined }>;
export interface FileFolderDemoSettings {
  fileFolderMyFiles?: boolean;
  fileFolderRecent?: boolean;
  fileFolderShared?: boolean;
  fileFolderQuickAccess?: boolean;
  fileFolderOrganization?: boolean;
  fileFolderSite?: boolean;
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
