import { nativeFormResource } from './localization/form-native.internal';
import { standardSchemaResolver } from '@hookform/resolvers/standard-schema';
import type { StandardSchemaV1 } from '@standard-schema/spec';
import { get, set } from 'react-hook-form';
import type { FieldErrors, FieldValues, RegisterOptions, Resolver, UseFormProps } from 'react-hook-form';
import type { FormUseFormProps } from './form.types';
import type { FormInputValidators } from './form-methods.internal';
import { assertFormValues } from './form-methods.internal';
/** Converts Standard Schema through the official resolver, preserving async transformations. */
export function createStandardSchemaResolver<TInput extends FieldValues, TContext, TOutput>(schema: StandardSchemaV1<TInput, TOutput>): Resolver<TInput, TContext, TOutput> {
    if (!schema || schema['~standard']?.version !== 1 || typeof schema['~standard'].validate !== 'function')
        throw new TypeError('Form expected a Standard Schema v1 validator.');
    return standardSchemaResolver<TInput, TContext, TOutput>(schema);
}
/** Removes absent initialization values without erasing RHF's own default options. */
export function resolveFormOptions<TInput extends FieldValues, TContext, TOutput>(props: FormUseFormProps<TInput, TContext, TOutput, boolean> | undefined, boundResolver?: Resolver<TInput, TContext, TOutput>): UseFormProps<TInput, TContext, TOutput> {
    const supplied = props ?? {};
    if (boundResolver && (supplied.schema !== undefined || supplied.resolver !== undefined))
        throw new TypeError('This Form suite already owns validation; its owner cannot be replaced.');
    if (supplied.schema !== undefined && supplied.resolver !== undefined)
        throw new TypeError('Form accepts either schema or resolver, not both.');
    assertFormValues(supplied.defaultValues);
    assertFormValues(supplied.values);
    const { schema, ...options } = supplied;
    delete options.strings;
    for (const key of Object.keys(options) as Array<keyof typeof options>)
        if (options[key] === undefined)
            delete options[key];
    const resolver = boundResolver ?? (schema !== undefined ? createStandardSchemaResolver<TInput, TContext, TOutput>(schema) : supplied.resolver);
    return resolver === undefined ? options : { ...options, resolver };
}
/** Generic fallback for unexpected failures; expected service errors remain consumer-owned. */
export const unexpectedFormFailureMessage = nativeFormResource.unexpectedFailure;
/** Retains the original exception for development diagnostics without a parallel error store. */
export function reportFormFailure(error: unknown): void {
    if (process.env.NODE_ENV !== 'production')
        console.error('Unexpected Form failure:', error);
}
/** Contains resolver exceptions in a non-root error that survives RHF's root-error clearing. */
export function withFormInputValidation<TInput extends FieldValues, TContext, TOutput>(resolver: Resolver<TInput, TContext, TOutput>, validators: FormInputValidators, getUnexpectedFailure: () => string = () => unexpectedFormFailureMessage): Resolver<TInput, TContext, TOutput> {
    return async (values, context, options) => {
        try {
            const result = await resolver(values, context, options);
            const errors: FieldErrors<TInput> = {};
            Object.assign(errors, result.errors);
            for (const validator of validators.values()) {
                const outcome = validator.validate(get(values, validator.name));
                if (outcome !== true)
                    set(errors, validator.name, { type: 'formInput', message: outcome });
            }
            return Object.keys(errors).length ? { values: {}, errors } : result;
        }
        catch (error) {
            reportFormFailure(error);
            // RHF removes errors.root before deciding whether onValid can run. Keep a real non-root blocker.
            const errors: FieldErrors<TInput> = {};
            set(errors, '__formValidation', { type: 'unexpected', message: getUnexpectedFailure() });
            return { values: {}, errors };
        }
    };
}
/** Contains custom validator failures at their boundary; expected validation results are unchanged. */
export function normalizeFormRules<TInput extends FieldValues>(rules: RegisterOptions<TInput> | undefined, internalValidator?: (value: unknown) => true | string, getUnexpectedFailure: () => string = () => unexpectedFormFailureMessage): RegisterOptions<TInput> | undefined {
    if (rules === undefined && internalValidator === undefined)
        return undefined;
    type Validator = NonNullable<RegisterOptions<TInput>['validate']>;
    type ValidateFunction = Extract<Validator, (...args: never[]) => unknown>;
    const safe = (validator: ValidateFunction): ValidateFunction => (async (value, values) => {
        try {
            return await validator(value, values);
        }
        catch (error) {
            reportFormFailure(error);
            return getUnexpectedFailure();
        }
    }) as ValidateFunction;
    const source = rules?.validate;
    if (typeof source === 'function') {
        const domainValidator = safe(source);
        if (internalValidator === undefined) return { ...rules, validate: domainValidator };
        const draftValidator = safe(internalValidator as ValidateFunction);
        // A function stays a function: RHF reports type/types.validate even after composition.
        const combined: ValidateFunction = async (value, values) => {
            const result = await domainValidator(value, values);
            return result === true || result === undefined ? draftValidator(value, values) : result;
        };
        return { ...rules, validate: combined };
    }
    const validators: Record<string, ValidateFunction> = Object.fromEntries(
        Object.entries(source ?? {}).map(([key, validator]) => [key, safe(validator)]),
    );
    if (internalValidator !== undefined) {
        // Consumer validator names remain intact, including formInput and this private base name.
        let internalKey = '__formInput';
        while (Object.prototype.hasOwnProperty.call(validators, internalKey)) internalKey += '_';
        validators[internalKey] = safe(internalValidator as ValidateFunction);
    }
    return { ...rules, ...(Object.keys(validators).length ? { validate: validators } : {}) } as RegisterOptions<TInput>;
}
/** Clones form records/arrays and Dates so editable fields cannot mutate submitted output.
 * Immutable browser File/Blob/FileList and other consumer class instances preserve identity.
 */
export function snapshotFormOutput<T>(value: T): T {
    const seen = new WeakMap<object, unknown>();
    function clone(current: unknown): unknown {
        if (current === undefined || typeof current !== 'object' || !current)
            return current;
        if (current instanceof Date)
            return new Date(current.getTime());
        const previous = seen.get(current);
        if (previous !== undefined)
            return previous;
        if (Array.isArray(current)) {
            const result: unknown[] = [];
            seen.set(current, result);
            for (const item of current)
                result.push(clone(item));
            return result;
        }
        if (Object.getPrototypeOf(current) !== Object.prototype)
            return current;
        const result: Record<string, unknown> = {};
        seen.set(current, result);
        for (const [key, item] of Object.entries(current))
            result[key] = clone(item);
        return result;
    }
    return clone(value) as T;
}
