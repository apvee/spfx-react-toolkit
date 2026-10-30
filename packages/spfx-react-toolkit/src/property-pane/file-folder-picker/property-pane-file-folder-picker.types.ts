import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { FileFolderPickerProps, FileFolderPickerValue } from '../../controls/file-folder-picker/file-folder-picker.types';
import type { JsonArrayPropertyKey } from '../property-pane-json-array-key.internal';

/** Safe direct property whose element domain exactly matches FileFolderPickerValue.
 * Optional and readonly arrays are accepted; scalar, legacy picker, untyped, mixed and nested properties are rejected.
 * Array values remain required even when multiselect is false. Prototype names and nested paths are never eligible.
 * @typeParam TProperties Web-part property bag containing a JSON file/folder reference array.
 * @example
 * type Props = { documents: FileFolderPickerValue[]; title: string };
 * const target: FileFolderPickerPropertyKey<Props> = 'documents';
 */
export type FileFolderPickerPropertyKey<TProperties extends object> = JsonArrayPropertyKey<TProperties, FileFolderPickerValue>;

type DistributiveOmit<T, Key extends PropertyKey> = T extends unknown ? Omit<T, Key> : never;

/** Native property-pane options for the hosted FileFolderPicker React control.
 * Inherits FileFolderPickerProps source restrictions, initialOptions display seeds, Field slots, callbacks other than
 * onChange and verified defaults, including the single responsive hosted dialog and current-user OneDrive discovery. The native field owns controlled state; defaultValue and onChange
 * are unavailable. Accepted explicit actions send fresh validated arrays through SPFx's onRender change callback;
 * hydration never writes properties. The host owns apply/persistence and existing delegated permissions.
 * Equivalent values/cosmetic changes preserve the pending nonreactive draft and mounted provider/session.
 * Source/policy or authoritative value changes cancel an open picker. Native disposal closes owned resources;
 * reopening creates a fresh metadata cache after the owner's final field lease is released.
 * @typeParam TProperties Web-part property bag containing an eligible direct reference array.
 * @example
 * const options: PropertyPaneFileFolderPickerProps<{ documents: FileFolderPickerValue[] }> = {
 *   instance: this, key: 'documents-picker', value: this.properties.documents,
 *   preset: 'focused', label: 'Documents', multiselect: true, extensions: ['.pdf'],
 * };
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
     * value equals the old property bag. Omission/equivalent revision preserves that draft. Does not remount the host.
     */
    readonly valueRevision?: string | number;
  };
