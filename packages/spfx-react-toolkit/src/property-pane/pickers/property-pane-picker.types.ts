/**
 * Native SPFx picker option and direct-array property contracts.
 *
 * @packageDocumentation
 * @docPage property-pane/pickers
 */
import type { JsonArrayPropertyKey } from '../property-pane-json-array-key.internal';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { PickerValue } from '../../controls/pickers/picker-values.types';

/**
 * Safe direct property names whose array element domain exactly matches the selected JSON reference type.
 * Optional and readonly arrays are accepted. Scalar, mixed, untyped, unknown and narrower element domains, nested paths and prototype-related names are rejected. The native factory still receives an array for single selection; pass [] for an absent optional property.
 * @typeParam TProperties - Owning web-part property bag.
 * @typeParam TValue - Complete matching toolkit picker JSON reference domain, not a UI option model.
 * @example Select a native property
 * ```ts
 * import type { PickerPropertyKey, SitePickerValue } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * type Properties = { sites?: readonly SitePickerValue[]; title: string };
 * const target: PickerPropertyKey<Properties, SitePickerValue> = 'sites';
 * ```
 */
export type PickerPropertyKey<TProperties extends object, TValue extends PickerValue> = JsonArrayPropertyKey<TProperties, TValue>;

/**
 * Remove React-owned selection members independently from each source branch, preserving source discriminants such as ContentTypePicker scope/list.
 * @typeParam T - Matching React control options or a union of source branches.
 * @typeParam K - Property names removed from every branch.
 */
type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never;

/**
 * Property Pane emits JSON value arrays, never raw UI models or only renderer keys.
 * The SPFx change callback receives accepted values; the host owns property updates and apply/persistence. The control does not persist values.
 * initialOptions and availableOptions stay distinct. Values removed from availableOptions remain selected.
 * Stable root/client/query runtime is preserved on prop updates, and original unresolved references are retained.
 * Available labels remain visible during hydration. Initial selections without display labels/models use AutoSkeleton only for the selection control region until initial hydration settles; later searches, refreshes and updates do not reactivate it.
 * Empty selections and mixed selections with known labels keep the normal control. Real terminal errors remain visible; no property is rewritten by hydration.
 * Inherits the React picker's accessible-name check: JSX/object labels require explicit
 * aria-label or aria-labelledby; otherwise supply a nonblank plain string label.
 * In default change mode, validateSelection can veto an action before a value is emitted.
 * Blur mode validates already committed ordinary gestures without reverting them; explicit ID
 * resolution still validates before emission. validationMessage overrides only displayed text.
 * Historical presentation combobox renders Dropdown with popup search; its trigger remains mounted through hydration.
 * Popup search and selection clear are separate. Blur validation runs after leaving the entire field and owned popup.
 * Selected tags stay in TagPicker rather than being duplicated in popup rows; colorfulTagAvatars affects tags only.
 * The native provider uses the initialized instance; callers do not install a second provider around the field. React `selectedOptions`, `defaultSelectedOptions`, `defaultPrincipals` and `onOptionSelect` are removed. Remaining React source, Field slots, callbacks, filtering, selection cardinality and defaults remain unchanged. Those callbacks retain their React semantics and do not acknowledge native persistence.
 * @typeParam TProperties - Owning web-part property bag with a compatible reference array.
 * @typeParam TControlProps - Matching React picker props, including any discriminated source restrictions.
 * @typeParam TValue - Complete JSON reference domain corresponding to TControlProps.
 * @remarks
 * ### Native pane theme and lifetime
 * All native picker fields, including the file/folder field, read tokens from the custom-field container in its owning document, independently of the web-part section theme. Usable opaque hex/comma-RGB surface and foreground anchors select a matching neutral light/dark base; remaining usable tokens overlay that base. An unavailable snapshot retains the same container's last valid theme, or neutral light before the first valid snapshot. Equivalent snapshots retain identity. Body-mounted menus/dialogs share this bridge. No network, persistent cache or polling is involved.
 * Native renders, host theme notifications, ancestor/reparenting changes and observed stylesheet replacement/text/load events refresh the snapshot without remounting. Unsignalled CSSOM-only changes need a host notification or another native render. Disposal releases observers, listeners and scheduled frames.
 * Same owner/key/DOM renders retain the React 17 root and provider; a changed owner/key replaces that session. Only the latest rendered factory may dispose it. Metadata cache leases are scoped to owner/document; closing the last field clears memory, so a later opening starts fresh. Separate roots can initialize clients independently. This does not guarantee tenant authorization or native Apply behavior.
 * @example Native options from an initialized host
 * ```ts
 * import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 * import type { PropertyPanePickerProps, SitePickerProps, SitePickerValue } from '@apvee/spfx-react-toolkit/property-pane/pickers';
 * type Properties = { sites?: SitePickerValue[] };
 * export function siteOptions(instance: BaseClientSideWebPart<Properties>, value: readonly SitePickerValue[]): PropertyPanePickerProps<Properties, SitePickerProps, SitePickerValue> {
 *   return { instance, key: 'sites', value, label: 'Sites' };
 * }
 * ```
 */
export type PropertyPanePickerProps<TProperties extends object, TControlProps, TValue extends PickerValue> =
  DistributiveOmit<TControlProps, 'selectedOptions' | 'defaultSelectedOptions' | 'defaultPrincipals' | 'onOptionSelect'> & {
    /** Required initialized web-part instance supplying context, clients and lifecycle ownership. Never serialized into the property value. */
  readonly instance: BaseClientSideWebPart<TProperties>;
    /** Required nonempty stable native field key within the owning web part. Never encode selection, query or time. */
    readonly key: string;
    /** Required complete ordered JSON references, even for single selection; pass [] for an undefined optional host property. Invalid updates preserve the last confirmed selection and disable mutation until valid values arrive. */
    readonly value: readonly TValue[];
    /** Explicitly adopt value and clear a pending nonreactive draft, even when its contents equal old props.
     * Initially omitted/undefined or unchanged revision preserves the draft when value is equivalent. Every revision change, including removal to undefined, adopts value without remounting or refreshing equivalent queries.
     */
    readonly valueRevision?: string | number;
  };
