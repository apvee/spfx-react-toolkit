import type { FieldProps } from '@fluentui/react-components';
import type { FileFolderPickerProps, FileFolderPickerScope, FileFolderPickerValue } from '../../packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker.types';
import type { JsonArrayPropertyKey } from '../../packages/spfx-react-toolkit/src/property-pane/property-pane-json-array-key.internal';
import type { SitePickerValue } from '../../packages/spfx-react-toolkit/src/controls/pickers/picker-values.types';
const defaults: FileFolderPickerProps = {};
const controlled: FileFolderPickerProps = { value: [], size: 'large', colorfulTagAvatars: true };
// @ts-expect-error The avatar flag is boolean, not a palette name.
const invalidAvatarColor: FileFolderPickerProps = { colorfulTagAvatars: 'colorful' };
void invalidAvatarColor;
const uncontrolled: FileFolderPickerProps = { defaultValue: [], scope: { kind: 'siteCollection' } };
declare const label: FieldProps['label'];
const nativeLabel: FileFolderPickerProps = { label };
// @ts-expect-error Controlled and uncontrolled state cannot be combined.
const both: FileFolderPickerProps = { value: [], defaultValue: [] };
// @ts-expect-error Small is unsupported.
const small: FileFolderPickerProps = { size: 'small' };
// @ts-expect-error Folder source requires a library.
const folder: FileFolderPickerScope = { kind: 'folder', webUrl: 'https://tenant.sharepoint.com', folder: 'A' };
// @ts-expect-error OneDrive source cannot carry SharePoint list metadata.
const drive: FileFolderPickerScope = { kind: 'oneDrive', baseUrl: 'https://tenant-my.sharepoint.com', list: 'Docs' };
// @ts-expect-error Unknown source kinds are rejected.
const source: FileFolderPickerScope = { kind: 'site' };
interface Props {
  exact: FileFolderPickerValue[];
  optional?: readonly FileFolderPickerValue[];
  wrong: SitePickerValue[];
  mixed: (FileFolderPickerValue | SitePickerValue)[];
  scalar: string;
  untyped: any;
  unknown: unknown;
  unknownArray: unknown[];
  anyArray: any[];
  neverArray: never[];
  nested: FileFolderPickerValue[][];
  'nested.files': FileFolderPickerValue[];
  'files[0]': FileFolderPickerValue[];
  '__proto__': FileFolderPickerValue[];
  constructor: FileFolderPickerValue[];
  prototype: FileFolderPickerValue[];
}
type Key = JsonArrayPropertyKey<Props, FileFolderPickerValue>;
const exact: Key = 'exact';
const optional: Key = 'optional';
type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;
const onlyExact: Equal<Key, 'exact' | 'optional'> = true;
void [defaults, controlled, uncontrolled, nativeLabel, both, small, folder, drive, source, exact, optional, onlyExact];

import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneFileFolderPicker } from '../../packages/spfx-react-toolkit/src/property-pane/file-folder-picker';
import type { FileFolderPickerPropertyKey, PropertyPaneFileFolderPickerProps } from '../../packages/spfx-react-toolkit/src/property-pane/file-folder-picker';
declare const instance: BaseClientSideWebPart<Props>;
const nativeProps: PropertyPaneFileFolderPickerProps<Props> = { instance, key: 'native-files', value: [], multiselect: false, colorfulTagAvatars: true };
const nativeKey: FileFolderPickerPropertyKey<Props> = 'exact';
PropertyPaneFileFolderPicker('exact', nativeProps);
PropertyPaneFileFolderPicker('optional', nativeProps);
// @ts-expect-error Native values are arrays even in single selection.
PropertyPaneFileFolderPicker('scalar', nativeProps);
// @ts-expect-error Legacy picker references are a distinct value domain.
PropertyPaneFileFolderPicker('wrong', nativeProps);
// @ts-expect-error Widened domains cannot be serialized safely.
PropertyPaneFileFolderPicker('mixed', nativeProps);
// @ts-expect-error Untyped properties are not exact value arrays.
PropertyPaneFileFolderPicker('anyArray', nativeProps);
// @ts-expect-error Native state has one owner; defaultValue is React-only.
const nativeDefault: PropertyPaneFileFolderPickerProps<Props> = { ...nativeProps, defaultValue: [] };
// @ts-expect-error Native callbacks come exclusively from SPFx onRender.
const nativeChange: PropertyPaneFileFolderPickerProps<Props> = { ...nativeProps, onChange: () => undefined };
void [nativeKey, nativeDefault, nativeChange];

const currentDrive: FileFolderPickerScope = { kind: 'oneDrive' };
// @ts-expect-error Only the current user's drive is supported.
const explicitDrive: FileFolderPickerScope = { kind: 'oneDrive', baseUrl: 'https://tenant-my.sharepoint.com/personal/other' };
// @ts-expect-error Hosted presentation is always dialog.
const popup: FileFolderPickerProps = { presentation: 'popup' };
void [currentDrive, explicitDrive, popup];

const flatConfiguration: FileFolderPickerProps = {
 title: 'Documents', hint: 'Consumer description', layout: 'tiles', stickyHeaders: true,
 initialSort: { fieldName: 'FileLeafRef', ascending: false }, initialFilter: { fieldName: 'Category', value: 'A' },
 visibleLocations: ['site', 'quickAccess'], navigationVisible: true, retainSelectionOnNavigate: true,
 graphEnabled: false, claimsChallengeEnabled: false, uploadEnabled: false, createFolderEnabled: false,
 fromLinkMode: 'navigation', localFilesMode: 'disabled', requiredAccess: 'read', trayPrompt: 'none', onReady() {},
};
// @ts-expect-error No unvalidated raw SDK options bag.
const rawConfiguration: FileFolderPickerProps = { pickerOptions: {} };
// @ts-expect-error Unverified local-file pick result contract is not exposed.
const localPick: FileFolderPickerProps = { localFilesMode: 'pick' };
// @ts-expect-error Theme derives from the host/provider.
const themeOverride: FileFolderPickerProps = { theme: 'dark' };
void [flatConfiguration, rawConfiguration, localPick, themeOverride];

import type { FileFolderPickerPreset } from '../../packages/spfx-react-toolkit/src/controls/file-folder-picker';
import type { FileFolderPickerConfigurationProps } from '../../packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker-config.internal';
const preset: FileFolderPickerPreset = 'focused';
const focused: FileFolderPickerProps = { preset, scope: { kind: 'library', webUrl: 'https://tenant.sharepoint.com/sites/hr', list: 'Documents' }, itemTypes: 'files', multiselect: true, extensions: ['.pdf', '.docx'], value: [], onChange: data => console.log(data.values) };
const extended: FileFolderPickerProps = { preset: 'extended', defaultValue: [], graphEnabled: undefined };
const nativeFocused: PropertyPaneFileFolderPickerProps<Props> = { instance, key: 'focused', value: [], preset: 'focused' };
PropertyPaneFileFolderPicker('exact', nativeFocused);
const configurationFocused: FileFolderPickerConfigurationProps = { preset: 'focused', sourceItems: [] };
// @ts-expect-error Unknown preset is rejected.
const unknownPreset: FileFolderPickerProps = { preset: 'custom' };
// @ts-expect-error Explicit legacy false cannot be silently reinterpreted by a preset.
const mixedGraph: FileFolderPickerProps = { preset: 'focused', graphEnabled: false };
// @ts-expect-error Explicit empty legacy navigation is also a mixed contract.
const mixedNavigation: FileFolderPickerProps = { preset: 'extended', visibleLocations: [] };
// @ts-expect-error Preset native options retain the same discriminator.
const nativeMixed: PropertyPaneFileFolderPickerProps<Props> = { instance, key: 'mixed', value: [], preset: 'focused', title: 'Legacy' };
// @ts-expect-error Configuration omission must preserve the discriminator.
const configurationMixed: FileFolderPickerConfigurationProps = { preset: 'focused', graphEnabled: false };
// @ts-expect-error Presets preserve controlled/uncontrolled exclusivity.
const presetBoth: FileFolderPickerProps = { preset: 'focused', value: [], defaultValue: [] };
// @ts-expect-error Presets do not make scalar web-part properties eligible.
PropertyPaneFileFolderPicker('scalar', nativeFocused);
void [focused, extended, configurationFocused, unknownPreset, mixedGraph, mixedNavigation, nativeMixed, configurationMixed, presetBoth];

import type { FileFolderPickerFileCategory } from '../../packages/spfx-react-toolkit/src/controls/file-folder-picker';
const category: FileFolderPickerFileCategory = 'images';
const categoryControl: FileFolderPickerProps = { preset: 'focused', fileCategory: category };
const categoryNative: PropertyPaneFileFolderPickerProps<Props> = { instance, key: 'category', value: [], fileCategory: 'media' };
// @ts-expect-error Categories are a bounded union.
const invalidCategory: FileFolderPickerProps = { fileCategory: 'stockImages' };
void [categoryControl, categoryNative, invalidCategory];

import type { FileFolderPickerStrings, FileFolderPickerStringsOverride } from '../../packages/spfx-react-toolkit/src/controls/file-folder-picker/file-folder-picker-strings.types';
declare const fullFileStrings: FileFolderPickerStrings;
const fileStrings: FileFolderPickerStringsOverride = { labels: { removeItem: name => `Remove ${name}` }, errors: { 'selection-policy': 'Invalid choice' } };
const localizedFile: FileFolderPickerProps = { preset: 'focused', strings: fileStrings };
const fullyLocalizedFile: FileFolderPickerProps = { strings: fullFileStrings };
declare const fileInstance: PropertyPaneFileFolderPickerProps<{ files: FileFolderPickerValue[] }>['instance'];
const localizedPaneFile: PropertyPaneFileFolderPickerProps<{ files: FileFolderPickerValue[] }> = { instance: fileInstance, key: 'files', value: [], strings: fileStrings };
// @ts-expect-error Strings cannot re-enable forbidden legacy preset labels.
const localizedLegacyTitle: FileFolderPickerProps = { preset: 'focused', strings: fileStrings, title: 'Legacy' };
// @ts-expect-error Safe error keys remain a closed set.
const invalidFileStrings: FileFolderPickerStringsOverride = { errors: { backend: 'Raw error' } };
void [localizedFile, fullyLocalizedFile, localizedPaneFile, localizedLegacyTitle, invalidFileStrings];
