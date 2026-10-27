import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';
import type { FieldArrayPath, FieldArrayPathValue, FieldPath, FieldPathValue, FieldValues, UseFieldArrayProps } from 'react-hook-form';
import type { ScopedFormSuite } from './form.types';
import type { FormHsvColor } from './form-value.types';

/** Removes optional absence while deriving nested object/array path compatibility. */
type Present<TValue> = Exclude<TValue, undefined>;
/** Atomic/native values and arrays excluded from nested record composition paths. */
type NativeOrArrayValue = Date | File | FileList | Blob | FormHsvColor | readonly unknown[];
/**
 * Nested record paths supported by Object/ObjectSection scopes. Native atomic values and arrays are excluded from object scopes.
 *
 * @example
 * ```tsx
 * const path: FormObjectPath<{ customer: { name: string } }> = 'customer';
 * ```
 */
export type FormObjectPath<TInput extends FieldValues> = {
    [TPath in FieldPath<TInput>]: Present<FieldPathValue<TInput, TPath>> extends FieldValues
        ? Present<FieldPathValue<TInput, TPath>> extends NativeOrArrayValue ? never : TPath
        : never;
}[FieldPath<TInput>];
/** Record model selected by an Object path.
 * @example `FormObjectPathValue<{ customer: { name: string } }, 'customer'>`
 */
export type FormObjectPathValue<TInput extends FieldValues, TPath extends FormObjectPath<TInput>> = Extract<Present<FieldPathValue<TInput, TPath>>, FieldValues>;
/**
 * Nested array paths supported by Array/ArraySection scopes.
 *
 * @example
 * ```tsx
 * const path: FormArrayPath<{ rows: { sku: string }[] }> = 'rows';
 * ```
 */
export type FormArrayPath<TInput extends FieldValues> = {
    [TPath in FieldArrayPath<TInput>]: FieldArrayPathValue<TInput, TPath>[number] extends FieldValues
        ? FieldArrayPathValue<TInput, TPath>[number] extends NativeOrArrayValue ? never : TPath
        : never;
}[FieldArrayPath<TInput>];
/** Complete live item model for a structured array path.
 * @example `FormArrayItemValue<{ rows: { sku: string }[] }, 'rows'>`
 */
export type FormArrayItemValue<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> = Extract<FieldArrayPathValue<TInput, TPath>[number], FieldValues>;
/** Relative field family and direct validation messages for an Object scope.
 * @example `({ Fields }) => <Fields.Input name="name" />`
 */
export interface FormObjectContext<TObject extends FieldValues> {
    /** Relative names are checked against the selected record. */
    Fields: ScopedFormSuite<TObject>;
    /** First direct object-level validation message, if present. */
    error?: string;
    /** All direct messages, deduplicated in deterministic order. */
    errors: readonly string[];
}
/** Headless scope for a nested record; adds no DOM layout.
 * @example `<Order.Object name="customer">{({ Fields }) => <Fields.Input name="name" />}</Order.Object>`
 */
export interface FormObjectProps<TInput extends FieldValues, TPath extends FormObjectPath<TInput>> {
    /** Relative path to the selected record. */
    name: TPath;
    /** Render nested fields, objects or structured arrays. */
    children: (context: FormObjectContext<FormObjectPathValue<TInput, TPath>>) => React.ReactNode;
}
/** Stable item identity and a field family that follows its current index.
 * @example `items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />)`
 */
export interface FormArrayItemContext<TItem extends FieldValues> {
    /** RHF-generated React identity, separate from the domain's id and array index. */
    key: string;
    /** Current zero-based position; changes when items are reordered. */
    index: number;
    /** Current number of records in the array. */
    total: number;
    /** Whether this record currently occupies the first position. */
    isFirst: boolean;
    /** Whether this record currently occupies the last position. */
    isLast: boolean;
    /** Stable relative family that reads this item's current absolute path on render. */
    Fields: ScopedFormSuite<TItem>;
}
/** Headless RHF record-array operations; operations never move native focus.
 * Invalid indices and duplicate/empty removal lists are no-ops. Disabled Roots/scopes
 * suppress every operation. update/replace use RHF's new-item/remount semantics.
 * @example `({ append }) => <button type="button" onClick={() => append({ sku: '' })}>Add</button>`
 */
export interface FormArrayContext<TItem extends FieldValues> {
    /** Current metadata; render records with item.key as their React key. */
    items: readonly FormArrayItemContext<TItem>[];
    /** Appends one complete record without focusing it. */
    append: (value: TItem) => void;
    /** Prepends one complete record without focusing it. */
    prepend: (value: TItem) => void;
    /** Inserts one record at an integer from zero through the array length. */
    insert: (index: number, value: TItem) => void;
    /** Removes valid distinct indices; omitted/undefined removes all records. */
    remove: (index?: number | number[]) => void;
    /** Replaces one complete record at a valid index, remounting its RHF fields. */
    update: (index: number, value: TItem) => void;
    /** Replaces all records, producing new RHF item identities. */
    replace: (values: TItem[]) => void;
    /** Moves a record between valid integer positions, preserving its identity. */
    move: (from: number, to: number) => void;
    /** Swaps two valid integer positions, preserving their identities. */
    swap: (first: number, second: number) => void;
    /** First direct array/root validation message, if present. */
    error?: string;
    /** Direct array and RHF array-root messages in deterministic deduplicated order. */
    errors: readonly string[];
}
/** Domain rules for the whole array; cannot accompany a schema/resolver owner.
 * @example `{ minLength: { value: 1, message: 'Add a row' } }`
 */
export type FormArrayRules<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> = NonNullable<UseFieldArrayProps<TInput, TPath>['rules']>;
/** Headless structured record array; primitive arrays belong to multi-value controls.
 * @example `<Order.Array name="rows">{({ items }) => items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />)}</Order.Array>`
 */
export interface FormArrayProps<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> {
    /** Relative structured-array path in the current record scope. */
    name: TPath;
    /** Rules-only validation for the array; rejected when a resolver owns validation. */
    rules?: FormArrayRules<TInput, TPath>;
    /** Render records and operation controls; the component adds no layout. */
    children: (context: FormArrayContext<FormArrayItemValue<TInput, TPath>>) => React.ReactNode;
}
/** Generic Object member retaining path-specific child model inference.
 * @example `const ObjectScope = Order.Object;`
 */
export type FormObjectComponent<TInput extends FieldValues> = <TPath extends FormObjectPath<TInput>>(props: FormObjectProps<TInput, TPath>) => React.ReactElement;
/** Generic Array member retaining path-specific item model inference.
 * @example `const ArrayScope = Order.Array;`
 */
export type FormArrayComponent<TInput extends FieldValues> = <TPath extends FormArrayPath<TInput>>(props: FormArrayProps<TInput, TPath>) => React.ReactElement;

/**
 * Accessible section naming: provide title, aria-label or aria-labelledby. Title is a literal-content slot in the owned legend.
 *
 * @example
 * ```ts
 * const naming: FormSectionName = { title: 'Customer' };
 * ```
 */
export type FormSectionName = {
    /** Literal phrasing content inside the native legend. Objects/callbacks configure an inner span and its ref. */
    title: ContentSlot<'span'>;
    /** Optional explicit accessible name overriding the legend. */
    'aria-label'?: string;
    /** Optional ID references overriding the legend. */
    'aria-labelledby'?: string;
} | {
    /** Optional literal legend content; object/callback refs identify its inner span. */
    title?: ContentSlot<'span'>;
    /** Explicit accessible name when a legend is omitted. */
    'aria-label': string;
    /** Optional external naming element IDs. */
    'aria-labelledby'?: string;
} | {
    /** Optional literal legend content; object/callback refs identify its inner span. */
    title?: ContentSlot<'span'>;
    /** Optional explicit accessible name. */
    'aria-label'?: string;
    /** External naming element IDs when a legend is omitted. */
    'aria-labelledby': string;
};
/** Shared native fieldset attributes and a reduced Fluent Card surface.
 * Caller children own custom layout; generated styles use toolkit classes.
 * @example `{ title: 'Customer', description: 'Contact details', appearance: 'outline' }`
 */
export type FormSectionProps = Omit<React.FieldsetHTMLAttributes<HTMLFieldSetElement>, 'name' | 'title' | 'children' | 'disabled' | 'dangerouslySetInnerHTML' | 'aria-label' | 'aria-labelledby'> & FormSectionName & {
    /** Literal description in the native hint; props/callbacks configure an inner div and its ref.
     * Empty static objects render their container without adding an empty aria-describedby reference. */
    description?: ContentSlot<'div'>;
    /** Adds to Root and enclosing scope restrictions and omits disabled adapter values on submission. */
    disabled?: boolean;
    /** Native Fluent Card appearance; defaults to filled. */
    appearance?: 'outline' | 'filled' | 'filled-alternative' | 'subtle';
};
/** Accessible visual nested-record scope with a fieldset/legend, description and direct errors.
 * No copied slots or layout configuration; compose custom layout inside children.
 * @example `<Order.ObjectSection name="customer" title="Customer">{({ Fields }) => <Fields.Input name="name" />}</Order.ObjectSection>`
 */
export type FormObjectSectionProps<TInput extends FieldValues, TPath extends FormObjectPath<TInput>> = FormSectionProps & FormObjectProps<TInput, TPath>;
/**
 * Localized visible actions, accessible action names and polite status announcements. Every callback receives the current item label; moved also receives one-based position and total.
 *
 * @example
 * ```ts
 * const strings: FormArraySectionStrings = {
 *   add: 'Aggiungi',
 *   empty: 'Nessuna riga',
 *   remove: label => `Elimina ${label}`
 * };
 * ```
 */
export interface FormArraySectionStrings {
    /** Add button text; defaults to Add item. */
    add?: string;
    /** Empty-array text; defaults to No items. */
    empty?: string;
    /** Visible and accessible Remove action text. */
    remove?: (label: string) => string;
    /** Visible and accessible Move up action text. */
    moveUp?: (label: string) => string;
    /** Visible and accessible Move down action text. */
    moveDown?: (label: string) => string;
    /** Announcement after adding an item. */
    added?: (label: string) => string;
    /** Announcement after removing an item. */
    removed?: (label: string) => string;
    /** Announcement after moving an item; position is one-based. */
    moved?: (label: string, position: number, total: number) => string;
}
/** Visual structured array with stable ordered items and constrained Add/Remove/Move actions.
 * Bounds constrain these UI actions, not validation, resets or external method writes.
 * Action focus skips unavailable targets; external resets never initiate focus movement.
 * @example `<Order.ArraySection name="rows" title="Rows" createItem={() => ({ sku: '' })} renderItem={({ Fields }) => <Fields.Input name="sku" />} />`
 */
export type FormArraySectionProps<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> = FormSectionProps & {
    /** Relative structured-record array path. Primitive arrays use multi-value adapters. */
    name: TPath;
    /** Creates one complete live-domain record for each allowed Add action. */
    createItem: () => FormArrayItemValue<TInput, TPath>;
    /** Renders caller-owned fields/layout using stable relative item metadata. */
    renderItem: (item: FormArrayItemContext<FormArrayItemValue<TInput, TPath>>) => React.ReactNode;
    /** Minimum UI item count; nonnegative integer, default zero. Does not create missing rows. */
    minItems?: number;
    /** Maximum UI item count; integer at least minItems, default unbounded. */
    maxItems?: number;
    /** Show Remove actions; defaults to true. */
    removable?: boolean;
    /** Show Move actions; defaults to false. Boundary actions are disabled. */
    reorderable?: boolean;
    /** Accessible list-item name used in actions and announcements; omission uses the per-form `arraySection.itemLabel` formatter with the one-based index. */
    itemLabel?: (item: FormArrayItemContext<FormArrayItemValue<TInput, TPath>>) => string;
    /** Localized action, empty and status text; omitted members inherit the per-form `arraySection` dictionary. */
    strings?: FormArraySectionStrings;
};
/** Generic ObjectSection member with path-specific relative child inference.
 * @example `const CustomerSection = Order.ObjectSection;`
 */
export type FormObjectSectionComponent<TInput extends FieldValues> = <TPath extends FormObjectPath<TInput>>(props: FormObjectSectionProps<TInput, TPath>) => React.ReactElement;
/** Generic ArraySection member with path-specific complete item inference.
 * @example `const RowsSection = Order.ArraySection;`
 */
export type FormArraySectionComponent<TInput extends FieldValues> = <TPath extends FormArrayPath<TInput>>(props: FormArraySectionProps<TInput, TPath>) => React.ReactElement;
