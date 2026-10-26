import type { JsonArrayPropertyKey } from '../property-pane-json-array-key.internal';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { PickerValue } from '../../controls/pickers/picker-values.types';

/** Direct safe property name whose array element domain exactly matches TValue. Rejects scalar, any, unknown, nested paths and prototype names.
 * @example type SiteKey = PickerPropertyKey<{ sites: SitePickerValue[]; title: string }, SitePickerValue>; // 'sites'
 */
export type PickerPropertyKey<TProperties extends object, TValue extends PickerValue> = JsonArrayPropertyKey<TProperties, TValue>;

// Omit must preserve discriminated source branches such as ContentTypePicker scope/list.
type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never;

/** Property Pane emits JSON value arrays, never raw UI models or only renderer keys.
 * SPFx changeCallback owns property updates and apply/persistence; the control itself does not store values.
 * initialOptions and availableOptions stay distinct. Values removed from availableOptions remain selected.
 * Stable root/client/query runtime is preserved on prop updates, and original unresolved references are retained.
 * Available labels remain visible during hydration. Initial selections without display labels/models use AutoSkeleton only for the input region until initial hydration settles; later searches, refreshes and updates do not reactivate it.
 * Empty selections and mixed selections with known labels keep the normal input. Real terminal errors remain visible; no property is rewritten by hydration.
 * Inherits the React picker's accessible-name check: JSX/object labels require explicit
 * aria-label or aria-labelledby; otherwise supply a nonblank plain string label.
 * validateSelection can veto an action before a value is emitted; validationMessage only
 * overrides its displayed message. Explicit ID resolution validates even in blur mode.
 * @example type Props = PropertyPanePickerProps<{sites:SitePickerValue[]}, SitePickerProps, SitePickerValue>;
 */
export type PropertyPanePickerProps<TProperties extends object, TControlProps, TValue extends PickerValue> =
  DistributiveOmit<TControlProps, 'selectedOptions' | 'defaultSelectedOptions' | 'defaultPrincipals' | 'onOptionSelect'> & {
    /** Required initialized web-part instance supplying context, clients and lifecycle ownership. Never serialized into the property value. */
  readonly instance: BaseClientSideWebPart<TProperties>;
    /** Required stable native field key within the owning web part. Never encode selection, query or time. */
    readonly key: string;
    /** Current JSON references; undefined optional host property should be passed as []. Authoritative selection. */
    readonly value: readonly TValue[];
    /** Explicitly adopt value and clear a pending nonreactive draft, even when its contents equal old props.
     * Omitted/equivalent revision preserves pending emission; changing it does not remount or refresh queries.
     */
    readonly valueRevision?: string | number;
  };
