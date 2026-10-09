import * as React from 'react';
import { get, useFieldArray, useFormState } from 'react-hook-form';
import type { FieldValues, RegisterOptions, UseFieldArrayProps } from 'react-hook-form';
import type { FormArrayItemContext, FormArrayItemValue, FormArrayPath, FormArrayProps } from '../form-composition.types';
import type { ScopedFormSuite } from '../form.types';
import { useFieldScopeDisabled, useResolvedFieldPath } from '../field-path.internal';
import { useFormRuntimeContext } from '../form-context.internal';
import { assertFormFieldName, commitFormResetArrayStructure, registerFormResetArrayPath, requireFormMethodsMetadata, synchronizeFormResetArrayOperation } from '../form-methods.internal';
import { normalizeFormRules } from '../form-validation.internal';
import { getFormErrorMessages } from '../error-messages.internal';
type ItemScopes = { keys: readonly string[]; fields: ReadonlyMap<string, ScopedFormSuite<FieldValues>> };
/** Headless RHF structured-record array with stable relative item families.
 * Primitive arrays use multi-value controls. Render with item.key; domain id is independent.
 * All operations honor additive disabled state and never move native focus.
 * @example `<Order.Array name="rows">{({ items, append }) => <>{items.map(({ key, Fields }) => <Fields.Input key={key} name="sku" />)}<button type="button" onClick={() => append({ sku: '' })}>Add</button></>}</Order.Array>`
 */
export function FormArray<TInput extends FieldValues, TPath extends FormArrayPath<TInput>>({ name, rules, children }: FormArrayProps<TInput, TPath>): React.ReactElement {
    type TItem = FormArrayItemValue<TInput, TPath>;
    const { createScopedFields, disabled: rootDisabled, methods } = useFormRuntimeContext();
    const scopeDisabled = useFieldScopeDisabled();
    const disabled = rootDisabled || scopeDisabled;
    const resolvedName = useResolvedFieldPath(String(name));
    assertFormFieldName(resolvedName);
    if (rules !== undefined && requireFormMethodsMetadata(methods).hasResolver)
        throw new TypeError('Form array rules cannot accompany a schema/resolver validation owner.');
    const normalizedRules = React.useMemo(() => normalizeFormRules(rules as RegisterOptions<FieldValues> | undefined) as UseFieldArrayProps<FieldValues>['rules'], [rules]);
    React.useLayoutEffect(() => registerFormResetArrayPath(methods, resolvedName), [methods, resolvedName]);
    const { fields, append, prepend, insert, remove, update, replace, move, swap } = useFieldArray({ control: methods.control, name: resolvedName, rules: normalizedRules });
    const { errors: formErrors } = useFormState({ control: methods.control, name: resolvedName });
    React.useEffect(() => {
        if (normalizedRules !== undefined) methods.register(resolvedName, normalizedRules as RegisterOptions<FieldValues>);
    }, [methods, normalizedRules, resolvedName]);
    React.useLayoutEffect(() => commitFormResetArrayStructure(methods, resolvedName), [fields, methods, resolvedName]);
    // These maps contain only item component/path identity; RHF owns all semantic state.
    const [ItemPathContext] = React.useState(() => React.createContext<ReadonlyMap<string, string>>(new Map()));
    const [scopes, setScopes] = React.useState<ItemScopes>({ keys: [], fields: new Map() });
    const keys = fields.map(field => field.id);
    let renderedScopes = scopes;
    if (keys.length !== scopes.keys.length || keys.some((key, index) => key !== scopes.keys[index])) {
        const nextFields = new Map<string, ScopedFormSuite<FieldValues>>();
        for (const key of keys) {
            const existing = scopes.fields.get(key);
            const useItemPrefix = (): string => {
                const prefix = React.useContext(ItemPathContext).get(key);
                if (prefix === undefined) throw new Error('Form.Array item Fields must render inside their owning array.');
                return prefix;
            };
            nextFields.set(key, existing ?? createScopedFields(useItemPrefix));
        }
        renderedScopes = { keys, fields: nextFields };
        setScopes(renderedScopes);
    }
    const paths = React.useMemo(() => new Map(fields.map((field, index) => [field.id, `${resolvedName}.${index}`])), [fields, resolvedName]);
    const total = fields.length;
    const items = fields.map((field, index): FormArrayItemContext<TItem> => {
        const Fields = renderedScopes.fields.get(field.id);
        if (Fields === undefined) throw new Error('Form.Array could not create an item field scope.');
        return { key: field.id, index, total, isFirst: index === 0, isLast: index === total - 1, Fields: Fields as unknown as ScopedFormSuite<TItem> };
    });
    const fieldError: unknown = get(formErrors, resolvedName);
    // RHF array errors may be arrays carrying an own root property.
    const rootError: unknown = fieldError && typeof fieldError === 'object' && 'root' in fieldError ? fieldError.root : undefined;
    const errors = getFormErrorMessages(fieldError, rootError);
    const validIndex = (index: number): boolean => Number.isInteger(index) && index >= 0 && index < total;
    return <ItemPathContext.Provider value={paths}>{children({
        items, errors, error: errors[0],
        append: value => {
            if (disabled) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'append', index: total });
            append(value, { shouldFocus: false });
        },
        prepend: value => {
            if (disabled) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'prepend' });
            prepend(value, { shouldFocus: false });
        },
        insert: (index, value) => {
            if (disabled || !Number.isInteger(index) || index < 0 || index > total) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'insert', index });
            insert(index, value, { shouldFocus: false });
        },
        remove: index => {
            if (disabled) return;
            if (index === undefined) {
                synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'remove' });
                remove();
                return;
            }
            const indices = Array.isArray(index) ? index : [index];
            const uniqueIndices = new Set(indices);
            if (indices.length === 0 || uniqueIndices.size !== indices.length || !indices.every(validIndex)) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'remove', indices: uniqueIndices });
            remove(index);
        },
        update: (index, value) => {
            if (disabled || !validIndex(index)) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'update', index });
            update(index, value);
        },
        replace: values => {
            if (disabled) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'replace', count: values.length });
            replace(values);
        },
        move: (from, to) => {
            if (disabled || !validIndex(from) || !validIndex(to)) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'move', from, to });
            move(from, to);
        },
        swap: (first, second) => {
            if (disabled || !validIndex(first) || !validIndex(second)) return;
            synchronizeFormResetArrayOperation(methods, resolvedName, { type: 'swap', first, second });
            swap(first, second);
        },
    })}</ItemPathContext.Provider>;
}
