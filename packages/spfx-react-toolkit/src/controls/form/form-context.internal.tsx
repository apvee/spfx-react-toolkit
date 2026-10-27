import * as React from 'react';
import type { FormStrings } from './form-strings.types';
import type { FieldValues } from 'react-hook-form';
import type { FormMethods, ScopedFormSuite } from './form.types';
/** Private provider data; no values/dirty/touched state is duplicated here. */
export interface FormRuntimeContextValue {
    methods: FormMethods<FieldValues, unknown, unknown>;
    disabled: boolean;
    strings: FormStrings;
    Fields: ScopedFormSuite<FieldValues>;
    createScopedFields: (usePrefix: () => string) => ScopedFormSuite<FieldValues>;
}
export const FormRuntimeContext = React.createContext<FormRuntimeContextValue | undefined>(undefined);
/** Requires an active native Form Root. */
export function useFormRuntimeContext(): FormRuntimeContextValue {
    const context = React.useContext(FormRuntimeContext);
    if (!context)
        throw new Error('Form controls must be rendered inside Form.Root.');
    return context;
}
