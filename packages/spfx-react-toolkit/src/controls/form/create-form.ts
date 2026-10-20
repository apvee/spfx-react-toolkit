import type { StandardSchemaV1 } from '@standard-schema/spec';
import type { FieldValues, Resolver } from 'react-hook-form';
import type { CreateForm, FormSuite } from './form.types';
import { createFormSuiteRuntime } from './form';
import { createStandardSchemaResolver } from './form-validation.internal';
/** Creates a stable typed Form family, optionally binding one validation owner.
 * Live methods read TInput; a schema/resolver may submit a distinct transformed TOutput.
 * Create the factory outside render. Input binds the entire string domain and has no codec.
 * @param configuration - Standard Schema v1, one schema/resolver option, or omitted for rules-only/Root validation.
 * @returns A family whose methods and fields share one ownership identity.
 * @example `const Order = createForm<{ reference: string }>();`
 */
export const createForm: CreateForm = (<TInput extends FieldValues, TOutput = TInput, TContext = unknown>(configuration?: StandardSchemaV1<TInput, TOutput> | {
    /** Standard Schema v1 owner; submitted output follows its transformation. */
    schema: StandardSchemaV1<TInput, TOutput>;
    /** Excluded when schema owns validation; supplying both is rejected. */
    resolver?: never;
} | {
    /** React Hook Form resolver owner for validation and output transformation. */
    resolver: Resolver<TInput, TContext, TOutput>;
    /** Excluded when resolver owns validation; supplying both is rejected. */
    schema?: never;
}): FormSuite<TInput, TContext, TOutput, boolean> => {
    if (configuration === undefined)
        return createFormSuiteRuntime<TInput, TContext, TOutput, false>();
    if ('~standard' in configuration)
        return createFormSuiteRuntime<TInput, TContext, TOutput, true>(createStandardSchemaResolver(configuration));
    if (configuration.schema !== undefined && configuration.resolver !== undefined)
        throw new TypeError('createForm accepts either schema or resolver, not both.');
    if (configuration.schema !== undefined)
        return createFormSuiteRuntime<TInput, TContext, TOutput, true>(createStandardSchemaResolver(configuration.schema));
    if (typeof configuration.resolver === 'function')
        return createFormSuiteRuntime<TInput, TContext, TOutput, true>(configuration.resolver);
    throw new TypeError('createForm requires a Standard Schema v1 object or schema/resolver configuration.');
}) as CreateForm;
