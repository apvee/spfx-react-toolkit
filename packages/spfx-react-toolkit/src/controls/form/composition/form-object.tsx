import * as React from 'react';
import { get, useFormState } from 'react-hook-form';
import type { FieldValues } from 'react-hook-form';
import type { FormObjectPath, FormObjectPathValue, FormObjectProps } from '../form-composition.types';
import type { ScopedFormSuite } from '../form.types';
import { FieldPathScope, useResolvedFieldPath } from '../field-path.internal';
import { useFormRuntimeContext } from '../form-context.internal';
import { assertFormFieldName } from '../form-methods.internal';
import { getFormErrorMessages } from '../error-messages.internal';
/** Headless nested-record scope: relative fields and composition share an absolute prefix.
 * No visual grouping or layout is introduced.
 * @example `<Order.Object name="customer">{({ Fields }) => <Fields.Input name="name" />}</Order.Object>`
 */
export function FormObject<TInput extends FieldValues, TPath extends FormObjectPath<TInput>>({ name, children }: FormObjectProps<TInput, TPath>): React.ReactElement {
    const { Fields, methods } = useFormRuntimeContext();
    const resolvedName = useResolvedFieldPath(String(name));
    assertFormFieldName(resolvedName);
    const { errors: formErrors } = useFormState({ control: methods.control, name: resolvedName });
    const errors = getFormErrorMessages(get(formErrors, resolvedName));
    return <FieldPathScope prefix={resolvedName}>{children({
        Fields: Fields as unknown as ScopedFormSuite<FormObjectPathValue<TInput, TPath>>,
        errors,
        error: errors[0],
    })}</FieldPathScope>;
}
