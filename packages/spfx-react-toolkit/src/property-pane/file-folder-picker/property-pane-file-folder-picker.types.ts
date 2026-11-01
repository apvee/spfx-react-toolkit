/**
 * Native SPFx file/folder field options and exact JSON array property keys.
 *
 * @packageDocumentation
 * @docPage property-pane/file-folder-picker
 */
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { FileFolderPickerProps, FileFolderPickerValue } from '../../controls/file-folder-picker/file-folder-picker.types';
import type { JsonArrayPropertyKey } from '../property-pane-json-array-key.internal';

/** Safe direct property whose element domain exactly matches FileFolderPickerValue.
 * Optional and readonly arrays are accepted; scalar, legacy picker, untyped, mixed and nested properties are rejected.
 * Array values remain required even when multiselect is false. Prototype names and nested paths are never eligible.
 * @typeParam TProperties - Web-part property bag containing a JSON file/folder reference array.
 * @example Select a file reference property
 * ```ts
 * import type { FileFolderPickerPropertyKey } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
 * import type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
 * type Props = { documents?: readonly FileFolderPickerValue[]; title: string };
 * const target: FileFolderPickerPropertyKey<Props> = 'documents';
 * ```
 */
export type FileFolderPickerPropertyKey<TProperties extends object> = JsonArrayPropertyKey<TProperties, FileFolderPickerValue>;

type DistributiveOmit<T, Key extends PropertyKey> = T extends unknown ? Omit<T, Key> : never;

/** Native property-pane options for the hosted FileFolderPicker React control.
 * Inherits FileFolderPickerProps source restrictions, initialOptions display seeds, Field slots, callbacks other than
 * onChange and verified defaults, including the single responsive hosted dialog and current-user OneDrive discovery. The native field owns controlled state; defaultValue and onChange
 * are unavailable. Accepted explicit actions send fresh validated arrays through SPFx's onRender change callback;
 * hydration never writes properties. The host owns apply/persistence and existing delegated permissions.
 * Native theme and owner/document lease behavior follows {@link PropertyPanePickerProps}; the hosted Microsoft iframe receives its light/dark SDK theme on opening and is not reskinned live.
 * Equivalent values/cosmetic changes preserve the pending nonreactive draft and mounted provider/session.
 * Source/policy or authoritative value changes cancel an open picker. Native disposal closes owned resources;
 * reopening creates a fresh metadata cache after the owner's final field lease is released.
 * @typeParam TProperties - Web-part property bag containing an eligible direct reference array.
 * @example Native file options from an initialized host
 * ```ts
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { PropertyPaneFileFolderPickerProps } from '@apvee/spfx-react-toolkit/property-pane/file-folder-picker';
 * import type { FileFolderPickerValue } from '@apvee/spfx-react-toolkit/controls/file-folder-picker';
 * type Properties = { documents?: FileFolderPickerValue[] };
 * export function documentOptions(instance: BaseClientSideWebPart<Properties>, value: readonly FileFolderPickerValue[]): PropertyPaneFileFolderPickerProps<Properties> {
 *   return { instance, key: 'documents', value,
 *     preset: 'focused', label: 'Documents', multiselect: true, extensions: ['.pdf'] };
 * }
 * ```
 */
export type PropertyPaneFileFolderPickerProps<TProperties extends object> =
  DistributiveOmit<FileFolderPickerProps, 'value' | 'defaultValue' | 'onChange'> & {
    /** Required initialized web-part instance; supplies provider context, clients and lifecycle ownership. Never serialized. */
    readonly instance: BaseClientSideWebPart<TProperties>;
    /** Required nonempty stable field key within this web part. Do not encode the selection, query or timestamp. */
    readonly key: string;
    /** Required complete ordered JSON array, including single-selection mode; pass [] for an omitted optional property.
     * Invalid runtime input preserves the last confirmed selection and disables mutation until valid input arrives.
     */
    readonly value: readonly FileFolderPickerValue[];
    /** Optional authoritative reset token. Changing it adopts value and clears any nonreactive draft even when
     * value equals the old property bag. Initially omitted/undefined or unchanged revision preserves the draft for equivalent values. Removing a prior revision to undefined is a change and adopts value too. Does not remount the host.
     */
    readonly valueRevision?: string | number;
  };
