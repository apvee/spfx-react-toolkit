/**
 * Typed Form families, RHF value ownership and native Fluent bindings.
 *
 * @packageDocumentation
 * @docPage controls/form
 */
import type * as React from 'react';
import type { ContentSlot } from '../content-slot.types.internal';
import type { FieldArrayPath, FieldArrayPathValue, FieldPath, FieldPathValue, FieldValues, UseFieldArrayProps } from 'react-hook-form';
import type { ScopedFormSuite } from './form.types';
import type { FormHsvColor } from './form-value.types';

/**
 * Removes optional absence while deriving nested object/array path compatibility.
 * @typeParam TValue - Complete domain read and written by the custom control.
 */
type Present<TValue> = Exclude<TValue, undefined>;
/** Atomic/native values and arrays excluded from nested record composition paths. */
type NativeOrArrayValue = Date | File | FileList | Blob | FormHsvColor | readonly unknown[];
/**
 * Nested record paths supported by Object/ObjectSection scopes. Native atomic values and arrays are excluded from object scopes.
 * @example
 * ```ts
 * import { FormObjectPath } from '@apvee/spfx-react-toolkit/controls/form';
 * const value: FormObjectPath<{ customer: { name: string } }> = 'customer';
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormObjectPath<TInput extends FieldValues> = {
    [TPath in FieldPath<TInput>]: Present<FieldPathValue<TInput, TPath>> extends FieldValues
        ? Present<FieldPathValue<TInput, TPath>> extends NativeOrArrayValue ? never : TPath
        : never;
}[FieldPath<TInput>];
/**
 * Record model selected by an Object path.
 * @example
 * ```ts
 * import { FormObjectPathValue } from '@apvee/spfx-react-toolkit/controls/form';
 * const value: FormObjectPathValue<{ customer: { name: string } }, 'customer'> = { name: '' };
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export type FormObjectPathValue<TInput extends FieldValues, TPath extends FormObjectPath<TInput>> = Extract<Present<FieldPathValue<TInput, TPath>>, FieldValues>;
/**
 * Nested array paths supported by Array/ArraySection scopes.
 * @example
 * ```ts
 * import { FormArrayPath } from '@apvee/spfx-react-toolkit/controls/form';
 * const value: FormArrayPath<{ rows: { sku: string }[] }> = 'rows';
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 */
export type FormArrayPath<TInput extends FieldValues> = {
    [TPath in FieldArrayPath<TInput>]: FieldArrayPathValue<TInput, TPath>[number] extends FieldValues
        ? FieldArrayPathValue<TInput, TPath>[number] extends NativeOrArrayValue ? never : TPath
        : never;
}[FieldArrayPath<TInput>];
/**
 * Complete live item model for a structured array path.
 * @example
 * ```ts
 * import { FormArrayItemValue } from '@apvee/spfx-react-toolkit/controls/form';
 * const value: FormArrayItemValue<{ rows: { sku: string }[] }, 'rows'> = { sku: '' };
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export type FormArrayItemValue<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> = Extract<FieldArrayPathValue<TInput, TPath>[number], FieldValues>;
/**
 * Relative field family and direct validation messages for an Object scope.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { FormObjectContext } from '@apvee/spfx-react-toolkit/controls/form';
 * export function renderCustomer({ Fields }: FormObjectContext<{ name: string }>): React.ReactElement {
 *   return <Fields.Input name="name" />;
 * }
 * ```
 * @typeParam TObject - Selected nested live record used for relative field names.
 */
export interface FormObjectContext<TObject extends FieldValues> {
    /** Relative names are checked against the selected record. */
    Fields: ScopedFormSuite<TObject>;
    /**
     * First direct object-level validation message, if present.
     * Undefined means no direct validation message is available.
     */
    error?: string;
    /** All direct messages, deduplicated in deterministic order. */
    errors: readonly string[];
}
/**
 * Headless scope for a nested record; adds no DOM layout.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormObjectProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ customer: { name: string } }>();
 * const props: FormObjectProps<{ customer: { name: string } }, 'customer'> = {
 *   name: 'customer', children: ({ Fields }) => <Fields.Input name="name" />,
 * };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ customer: { name: '' } }} onSubmit={values => { console.log(values); }}>
 *     <Order.Object {...props} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export interface FormObjectProps<TInput extends FieldValues, TPath extends FormObjectPath<TInput>> {
    /** Relative path to the selected record. */
    name: TPath;
    /**
     * Render nested fields, objects or structured arrays.
     * @param context - Relative Fields and direct validation error/messages for the selected record, supplied during rendering; the callback must be pure.
     * @returns Caller-owned React content rendered within the current scope; no extra layout is inferred.
     */
    children: (context: FormObjectContext<FormObjectPathValue<TInput, TPath>>) => React.ReactNode;
}
/**
 * Stable item identity and a field family that follows its current index.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { FormArrayItemContext } from '@apvee/spfx-react-toolkit/controls/form';
 * export function renderRows(context: FormArrayItemContext<{ sku: string }>): React.ReactElement {
 *   return <context.Fields.Input name="sku" />;
 * }
 * ```
 * @typeParam TItem - Complete live structured-array item record.
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
/**
 * Headless RHF record-array operations; operations never move native focus.
 * Invalid indices and duplicate/empty removal lists are no-ops. Disabled Roots/scopes
 * suppress every operation. update/replace use RHF's new-item/remount semantics.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { FormArrayContext } from '@apvee/spfx-react-toolkit/controls/form';
 * export function renderRows(context: FormArrayContext<{ sku: string }>): React.ReactElement {
 *   return <><button type="button" onClick={() => context.append({ sku: "" })}>Add</button>{context.items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />)}</>;
 * }
 * ```
 * @typeParam TItem - Complete live structured-array item record.
 */
export interface FormArrayContext<TItem extends FieldValues> {
    /** Current metadata; render records with item.key as their React key. */
    items: readonly FormArrayItemContext<TItem>[];
    /**
     * Appends one complete record without focusing it.
     * @param value - Complete live-domain value to write or convert; no implicit domain coercion occurs.
     */
    append: (value: TItem) => void;
    /**
     * Prepends one complete record without focusing it.
     * @param value - Complete live-domain value to write or convert; no implicit domain coercion occurs.
     */
    prepend: (value: TItem) => void;
    /**
     * Inserts one record at an integer from zero through the array length.
     * @param index - Required integer insertion position from zero through the current item count, inclusive. Invalid positions are a no-op.
     * @param value - One complete new record inserted at that position; it is not merged with another record.
     */
    insert: (index: number, value: TItem) => void;
    /**
     * Removes valid distinct indices; omitted/undefined removes all records.
     * @param index - One valid integer item index or a nonempty array of distinct valid integer indices. Omitted/undefined removes all records; an invalid index, duplicate or empty list makes the entire request a no-op.
     */
    remove: (index?: number | number[]) => void;
    /**
     * Replaces one complete record at a valid index, remounting its RHF fields.
     * @param index - Required integer position of an existing item, from zero through item count minus one. Invalid positions are a no-op.
     * @param value - Complete replacement record for that item; RHF remounts the updated fields with a new item identity.
     */
    update: (index: number, value: TItem) => void;
    /**
     * Replaces all records, producing new RHF item identities.
     * @param values - Complete replacement array; no existing record is partially merged and RHF generates new item identities.
     */
    replace: (values: TItem[]) => void;
    /**
     * Moves a record between valid integer positions, preserving its identity.
     * @param from - Existing zero-based source position; invalid positions are ignored.
     * @param to - Existing zero-based destination position; invalid positions are ignored.
     */
    move: (from: number, to: number) => void;
    /**
     * Swaps two valid integer positions, preserving their identities.
     * @param first - First valid zero-based item position.
     * @param second - Second valid zero-based item position.
     */
    swap: (first: number, second: number) => void;
    /**
     * First direct array/root validation message, if present.
     * Undefined means no direct validation message is available.
     */
    error?: string;
    /** Direct array and RHF array-root messages in deterministic deduplicated order. */
    errors: readonly string[];
}
/**
 * Domain rules for the whole array; cannot accompany a schema/resolver owner.
 * @example
 * ```ts
 * import { FormArrayRules } from '@apvee/spfx-react-toolkit/controls/form';
 * const rules: FormArrayRules<{ rows: { sku: string }[] }, 'rows'> = { minLength: { value: 1, message: 'Add a row' } };
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export type FormArrayRules<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> = NonNullable<UseFieldArrayProps<TInput, TPath>['rules']>;
/**
 * Headless structured record array; primitive arrays belong to multi-value controls.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormArrayProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ rows: { sku: string }[] }>();
 * const props: FormArrayProps<{ rows: { sku: string }[] }, 'rows'> = {
 *   name: 'rows', children: ({ items }) => items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />),
 * };
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ rows: [] }} onSubmit={values => { console.log(values); }}>
 *     <Order.Array {...props} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export interface FormArrayProps<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> {
    /** Relative structured-array path in the current record scope. */
    name: TPath;
    /**
     * Rules-only validation for the array; rejected when a resolver owns validation.
     * When omitted, no field or array domain rules are registered here; schema/resolver ownership remains unchanged.
     */
    rules?: FormArrayRules<TInput, TPath>;
    /**
     * Render records and operation controls; the component adds no layout.
     * @param context - Current stable item metadata, array operations and direct errors, supplied during rendering; the callback must be pure.
     * @returns Caller-owned React content rendered within the current scope; no extra layout is inferred.
     */
    children: (context: FormArrayContext<FormArrayItemValue<TInput, TPath>>) => React.ReactNode;
}
/**
 * Generic Object member retaining path-specific child model inference.
 * @example
 * ```ts
 * import { createForm, FormObjectComponent } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ customer: { name: string } }>();
 * const Customer: FormObjectComponent<{ customer: { name: string } }> = Order.Object;
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 * @param props - Required relative record name and children render callback. Render inside the same suite’s Root; this headless scope accepts no native attributes or ref.
 * @returns The children render result under a relative record scope, without adding layout DOM or forwarding a ref.
 */
export type FormObjectComponent<TInput extends FieldValues> = <TPath extends FormObjectPath<TInput>>(props: FormObjectProps<TInput, TPath>) => React.ReactElement;
/**
 * Generic Array member retaining path-specific item model inference.
 * @example
 * ```ts
 * import { createForm, FormArrayComponent } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ rows: { sku: string }[] }>();
 * const Rows: FormArrayComponent<{ rows: { sku: string }[] }> = Order.Array;
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 * @param props - Required relative structured-array name and children render callback, plus optional array rules when no schema/resolver owns validation. Requires the same suite’s Root; no native attributes or ref are accepted.
 * @returns The children render result with stable item metadata and array operations, without adding layout DOM or forwarding a ref.
 */
export type FormArrayComponent<TInput extends FieldValues> = <TPath extends FormArrayPath<TInput>>(props: FormArrayProps<TInput, TPath>) => React.ReactElement;

/**
 * Accessible section naming: provide title, aria-label or aria-labelledby. Title is a literal-content slot in the owned legend.
 *
 * @example
 * ```ts
 * import { FormSectionName } from '@apvee/spfx-react-toolkit/controls/form';
 * const naming: FormSectionName = { title: 'Customer' };
 * ```
 */
export type FormSectionName = {
    /** Literal phrasing content inside the native legend. Objects/callbacks configure an inner span and its ref. */
    title: ContentSlot<'span'>;
    /**
     * Optional explicit accessible name overriding the legend.
     * When omitted, the native legend names the section unless aria-labelledby overrides it.
     */
    'aria-label'?: string;
    /**
     * Optional ID references overriding the legend.
     * When omitted, aria-label supplies the name when present; otherwise the native legend names the section.
     */
    'aria-labelledby'?: string;
} | {
    /**
     * Optional literal legend content; object/callback refs identify its inner span.
     * When omitted, supply an explicit outer accessible name; no visible legend is added.
     */
    title?: ContentSlot<'span'>;
    /** Explicit accessible name when a legend is omitted. */
    'aria-label': string;
    /**
     * Optional external naming element IDs.
     * When omitted, the required aria-label names the section.
     */
    'aria-labelledby'?: string;
} | {
    /**
     * Optional literal legend content; object/callback refs identify its inner span.
     * When omitted, supply an explicit outer accessible name; no visible legend is added.
     */
    title?: ContentSlot<'span'>;
    /**
     * Optional explicit accessible name.
     * When omitted, the required aria-labelledby references name the section.
     */
    'aria-label'?: string;
    /** External naming element IDs when a legend is omitted. */
    'aria-labelledby': string;
};
/**
 * Shared native fieldset attributes and a reduced Fluent Card surface.
 * Caller children own custom layout; generated styles use toolkit classes.
 * @example
 * ```ts
 * import { FormSectionProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const section: FormSectionProps = { title: 'Customer', description: 'Contact details', appearance: 'outline' };
 * ```
 */
export type FormSectionProps = Omit<React.FieldsetHTMLAttributes<HTMLFieldSetElement>, 'name' | 'title' | 'children' | 'disabled' | 'dangerouslySetInnerHTML' | 'aria-label' | 'aria-labelledby'> & FormSectionName & {
    /**
     * Literal description in the native hint; props/callbacks configure an inner div and its ref.
     * Empty static objects render their container without adding an empty aria-describedby reference.
     * When omitted, no additional description content or association is requested.
     */
    description?: ContentSlot<'div'>;
    /**
     * Adds to Root and enclosing scope restrictions and omits disabled adapter values on submission.
     * When omitted or false, adds no disabled restriction; enclosing Root/scope restrictions still apply.
     */
    disabled?: boolean;
    /**
     * Native Fluent Card appearance; defaults to filled.
     * When omitted, the native Fluent primitive appearance applies.
     * @defaultValue filled
     */
    appearance?: 'outline' | 'filled' | 'filled-alternative' | 'subtle';
};
/**
 * Accessible visual nested-record scope with a fieldset/legend, description and direct errors.
 * No copied slots or layout configuration; compose custom layout inside children.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormObjectSectionProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ customer: { name: string } }>();
 * const props = {
 *   name: 'customer', title: 'Customer', children: ({ Fields }) => <Fields.Input name="name" />,
 * } satisfies FormObjectSectionProps<{ customer: { name: string } }, 'customer'>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ customer: { name: '' } }} onSubmit={values => { console.log(values); }}>
 *     <Order.ObjectSection {...props} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export type FormObjectSectionProps<TInput extends FieldValues, TPath extends FormObjectPath<TInput>> = FormSectionProps & FormObjectProps<TInput, TPath>;
/**
 * Localized visible actions, accessible action names and polite status announcements. Every callback receives the current item label; moved also receives one-based position and total.
 *
 * @example
 * ```ts
 * import { FormArraySectionStrings } from '@apvee/spfx-react-toolkit/controls/form';
 * const strings: FormArraySectionStrings = {
 *   add: 'Aggiungi',
 *   empty: 'Nessuna riga',
 *   remove: label => `Elimina ${label}`
 * };
 * ```
 */
export interface FormArraySectionStrings {
    /**
     * Add button text; defaults to Add item.
     * When omitted, the per-form native arraySection.add label is used.
     */
    add?: string;
    /**
     * Empty-array text; defaults to No items.
     * When omitted, the per-form native arraySection.empty text is used.
     */
    empty?: string;
    /**
     * Visible and accessible Remove action text.
     * When omitted, the per-form native remove formatter is used.
     * @param label - Current caller item label, interpolated literally into visible or announced text.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    remove?: (label: string) => string;
    /**
     * Visible and accessible Move up action text.
     * When omitted, the per-form native moveUp formatter is used.
     * @param label - Current caller item label, interpolated literally into visible or announced text.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    moveUp?: (label: string) => string;
    /**
     * Visible and accessible Move down action text.
     * When omitted, the per-form native moveDown formatter is used.
     * @param label - Current caller item label, interpolated literally into visible or announced text.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    moveDown?: (label: string) => string;
    /**
     * Announcement after adding an item.
     * When omitted, the per-form native added formatter is used.
     * @param label - Current caller item label, interpolated literally into visible or announced text.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    added?: (label: string) => string;
    /**
     * Announcement after removing an item.
     * When omitted, the per-form native removed formatter is used.
     * @param label - Current caller item label, interpolated literally into visible or announced text.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    removed?: (label: string) => string;
    /**
     * Announcement after moving an item; position is one-based.
     * When omitted, the per-form native moved formatter is used.
     * @param label - Current caller item label, interpolated literally into visible or announced text.
     * @param position - One-based position after the action.
     * @param total - Current total item count after the action.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    moved?: (label: string, position: number, total: number) => string;
}
/**
 * Visual structured array with stable ordered items and constrained Add/Remove/Move actions.
 * Bounds constrain these UI actions, not validation, resets or external method writes.
 * Action focus skips unavailable targets; external resets never initiate focus movement.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { createForm, FormArraySectionProps } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ rows: { sku: string }[] }>();
 * const props = {
 *   name: 'rows', title: 'Rows', createItem: () => ({ sku: '' }), renderItem: ({ Fields }) => <Fields.Input name="sku" />,
 * } satisfies FormArraySectionProps<{ rows: { sku: string }[] }, 'rows'>;
 * export function Editor(): React.ReactElement {
 *   return <Order.Root defaultValues={{ rows: [] }} onSubmit={values => { console.log(values); }}>
 *     <Order.ArraySection {...props} />
 *   </Order.Root>;
 * }
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 */
export type FormArraySectionProps<TInput extends FieldValues, TPath extends FormArrayPath<TInput>> = FormSectionProps & {
    /** Relative structured-record array path. Primitive arrays use multi-value adapters. */
    name: TPath;
    /**
     * Creates one complete live-domain record for each allowed Add action.
     * @returns One complete live-domain record; the allowed Add action appends it.
     */
    createItem: () => FormArrayItemValue<TInput, TPath>;
    /**
     * Renders caller-owned fields/layout using stable relative item metadata.
     * @param item - Current stable item metadata and relative Fields; render callbacks must be pure.
     * @returns Caller-owned React content rendered within the current scope; no extra layout is inferred.
     */
    renderItem: (item: FormArrayItemContext<FormArrayItemValue<TInput, TPath>>) => React.ReactNode;
    /**
     * Minimum UI item count; nonnegative integer, default zero. Does not create missing rows.
     * When omitted, the minimum is zero; it constrains UI removal only and creates no records.
     * @defaultValue 0
     */
    minItems?: number;
    /**
     * Maximum UI item count; integer at least minItems, default unbounded.
     * When omitted, UI item count has no upper bound.
     */
    maxItems?: number;
    /**
     * Show Remove actions; defaults to true.
     * When omitted, Remove actions are shown.
     * @defaultValue true
     */
    removable?: boolean;
    /**
     * Show Move actions; defaults to false. Boundary actions are disabled.
     * When omitted, Move actions are hidden.
     * @defaultValue false
     */
    reorderable?: boolean;
    /**
     * Accessible list-item name used in actions and announcements; omission uses the per-form `arraySection.itemLabel` formatter with the one-based index.
     * When omitted, the per-form arraySection.itemLabel formats the one-based index.
     * @param item - Current stable item metadata and relative Fields; render callbacks must be pure.
     * @returns Localized visible/action/announcement text; interpolation retains caller text literally.
     */
    itemLabel?: (item: FormArrayItemContext<FormArrayItemValue<TInput, TPath>>) => string;
    /** Localized action, empty and status text; omitted members inherit the per-form `arraySection` dictionary. */
    strings?: FormArraySectionStrings;
};
/**
 * Generic ObjectSection member with path-specific relative child inference.
 * @example
 * ```ts
 * import { createForm, FormObjectSectionComponent } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ customer: { name: string } }>();
 * const Customer: FormObjectSectionComponent<{ customer: { name: string } }> = Order.ObjectSection;
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 * @param props - Required relative record name, children renderer and accessible section name through title, aria-label or aria-labelledby. Optional props configure the fieldset/surface; requires the same suite’s Root and does not accept a ref.
 * @returns A named native fieldset containing a Fluent Card, description, direct errors and caller-rendered relative fields; no ref is forwarded.
 */
export type FormObjectSectionComponent<TInput extends FieldValues> = <TPath extends FormObjectPath<TInput>>(props: FormObjectSectionProps<TInput, TPath>) => React.ReactElement;
/**
 * Generic ArraySection member with path-specific complete item inference.
 * @example
 * ```ts
 * import { createForm, FormArraySectionComponent } from '@apvee/spfx-react-toolkit/controls/form';
 * const Order = createForm<{ rows: { sku: string }[] }>();
 * const Rows: FormArraySectionComponent<{ rows: { sku: string }[] }> = Order.ArraySection;
 * ```
 * @typeParam TInput - Live form record used by field paths, getValues/watch and writes.
 * @typeParam TPath - Selected nested object/structured-array path, retaining its exact child model.
 * @param props - Required relative structured-array name, createItem, renderItem and accessible section name through title, aria-label or aria-labelledby. Optional props configure UI bounds/actions and the fieldset/surface; requires the same suite’s Root and does not accept a ref.
 * @returns A named native fieldset containing a Fluent Card, ordered items, localized actions and announcements; no ref is forwarded.
 */
export type FormArraySectionComponent<TInput extends FieldValues> = <TPath extends FormArrayPath<TInput>>(props: FormArraySectionProps<TInput, TPath>) => React.ReactElement;
