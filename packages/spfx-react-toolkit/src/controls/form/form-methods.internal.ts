import * as React from 'react';
import type { FormStrings } from './form-strings.types';
import type { FieldValues, RegisterOptions, UseFormReturn, UseFormProps } from 'react-hook-form';
import type { normalizeFormRules } from './form-validation.internal';
import type { FormMethods } from './form.types';
/** One reset notification; names and owner identities belong to this form only. */
export interface FormResetEvent {
    version: number;
    name?: string;
    preserveValues: boolean;
    preservedOwners: ReadonlySet<symbol>;
    source: 'explicit' | 'reactive' | 'async';
    /** Stable recipients captured before an array render can change mounted paths. */
    owners?: ReadonlySet<symbol>;
}
/** Identity-only structural changes; RHF still performs every value/field-array action. */
export type FormResetArrayOperation =
    | { type: 'append'; index: number }
    | { type: 'prepend' }
    | { type: 'insert'; index: number }
    | { type: 'remove'; indices?: ReadonlySet<number> | readonly number[] }
    | { type: 'update'; index: number }
    | { type: 'replace'; count: number }
    | { type: 'move'; from: number; to: number }
    | { type: 'swap'; first: number; second: number };
/** Private incomplete-draft validation keyed by stable field owner. */
export type FormInputValidators = Map<symbol, {
    name: string;
    validate: (value: unknown) => true | string;
}>;
/** Suite identity and lifecycle metadata; no form values are stored here. */
export interface FormMethodsMetadata {
    suiteIdentity: symbol;
    /** Latest committed dictionary owned by the useForm instance. */
    getStrings: () => FormStrings;
    /** Mounted Roots observe only committed dictionary changes. */
    stringsListeners: Set<() => void>;
    hasResolver: boolean;
    isResetting: boolean;
    /** Synchronous public writes can broadcast unnamed values without resetting RHF. */
    valueWriteDepth: number;
    inputValidators: FormInputValidators;
    resetOwners: Map<symbol, string>;
    mountedResetOwners: Map<symbol, string>;
    resetOwnerRegistrations: Map<symbol, symbol>;
    resetArrays: Map<symbol, string>;
    pendingOwnerResets: Map<symbol, FormResetEvent[]>;
    resetTextReaders: Map<symbol, () => string | undefined>;
    resetTextSnapshots: Map<string, { text: string; preservedOwners: Set<symbol> }>;
    fileOwners: Map<symbol, { read: () => FileList | undefined; clear: () => void }>;
    automaticReset?: { preservedOwners: Set<symbol>; preserveValues: boolean; source: 'reactive' | 'async' };
    committedResetOptions?: UseFormProps<FieldValues>['resetOptions'];
    clearFileValue: (name: string) => void;
    resetEvent: FormResetEvent;
    listeners: Set<(event: FormResetEvent) => void>;
}
const metadataByMethods = new WeakMap<object, FormMethodsMetadata>();
/** Excludes public mutation broadcasts from automatic-reset acknowledgement, including nested writes. */
function withFormValueWrite<TResult>(metadata: FormMethodsMetadata, write: () => TResult): TResult {
    metadata.valueWriteDepth++;
    try { return write(); }
    finally { metadata.valueWriteDepth--; }
}
// Browser file controls require the explicit nullable live domain when selection is absent.
// eslint-disable-next-line @rushstack/no-new-null
const clearedFileSelection = null;
/** Normalizes only registered package file inputs against their actual native selection. */
function synchronizeFormFileOwners(methods: object, target?: string, preservedOwners?: ReadonlySet<symbol>, keepValues = false, initialize = false): void {
    const metadata = requireFormMethodsMetadata(methods);
    const form = methods as Pick<UseFormReturn<FieldValues>, 'getValues'>;
    for (const [owner, input] of metadata.fileOwners) {
        const name = metadata.resetOwners.get(owner);
        if (name === undefined || (target !== undefined && name !== target && !name.startsWith(`${target}.`))) continue;
        const current = input.read();
        const live: unknown = form.getValues(name);
        if ((keepValues || preservedOwners?.has(owner)) && current?.length && live === current) continue;
        input.clear();
        if (live !== clearedFileSelection && (live !== undefined || initialize)) metadata.clearFileValue(name);
    }
}
/** Owns a real native file boundary; registration cannot restore external selections. */
export function registerFormFileInputOwner(methods: object, owner: symbol, input: { read: () => FileList | undefined; clear: () => void }): () => void {
    const metadata = requireFormMethodsMetadata(methods);
    metadata.fileOwners.set(owner, input);
    // Matching native/model identity is the only selection registration may retain.
    const name = metadata.resetOwners.get(owner);
    if (name !== undefined) synchronizeFormFileOwners(methods, name, new Set([owner]), false, true);
    return () => { if (metadata.fileOwners.get(owner) === input) metadata.fileOwners.delete(owner); };
}
/** Captures current dirty identities and optional query text immediately before one real reset. */
function prepareFormReset(methods: object, options?: { keepDirtyValues?: boolean; keepValues?: boolean }): Set<symbol> {
    const metadata = requireFormMethodsMetadata(methods);
    const form = methods as Pick<UseFormReturn<FieldValues>, 'getFieldState'>;
    const preservedOwners = new Set<symbol>();
    metadata.resetTextSnapshots.clear();
    if (options?.keepDirtyValues)
        for (const [owner, name] of metadata.resetOwners)
            if (form.getFieldState(name).isDirty) {
                preservedOwners.add(owner);
                const read = metadata.resetTextReaders.get(owner);
                if (!options.keepValues && read && [...metadata.resetArrays.values()].some(array => name.startsWith(`${array}.`))) {
                    const text = read();
                    if (text !== undefined) metadata.resetTextSnapshots.set(name, { text, preservedOwners });
                }
            }
    return preservedOwners;
}
/** Promise fulfillment boundary: prepare before RHF broadcasts regenerated array rows. */
export interface FormResetDefaultsBridge<TInput extends FieldValues, TContext, TOutput> {
    defaultValues: UseFormProps<TInput, TContext, TOutput>['defaultValues'];
    /** Public RHF configuration object; its contents are published only at React commit. */
    resetOptions: NonNullable<UseFormProps<TInput, TContext, TOutput>['resetOptions']>;
    commit: (methods: UseFormReturn<TInput, TContext, TOutput>, options: UseFormProps<TInput, TContext, TOutput>['resetOptions']) => void;
}
/** Preserves the consumer producer/rejection semantics while using only committed options. */
export function useFormResetDefaultsBridge<TInput extends FieldValues, TContext, TOutput>(defaultValues: UseFormProps<TInput, TContext, TOutput>['defaultValues']): FormResetDefaultsBridge<TInput, TContext, TOutput> {
    const committed = React.useRef<{ methods: UseFormReturn<TInput, TContext, TOutput>; options: UseFormProps<TInput, TContext, TOutput>['resetOptions'] }>();
    const resetOptions = React.useMemo<NonNullable<UseFormProps<TInput, TContext, TOutput>['resetOptions']>>(() => ({}), []);
    const wrapped = React.useMemo(() => typeof defaultValues !== 'function' ? defaultValues : async () => {
        const values = await defaultValues();
        const current = committed.current;
        if (current) requireFormMethodsMetadata(current.methods).automaticReset = { preservedOwners: prepareFormReset(current.methods, current.options), preserveValues: current.options?.keepValues === true, source: 'async' };
        return values;
    }, [defaultValues]);
    const commit = React.useCallback((methods: UseFormReturn<TInput, TContext, TOutput>, options: UseFormProps<TInput, TContext, TOutput>['resetOptions']) => {
        for (const key of Object.keys(resetOptions) as Array<keyof typeof resetOptions>) delete resetOptions[key];
        Object.assign(resetOptions, options);
        committed.current = { methods, options: resetOptions };
    }, [resetOptions]);
    return { defaultValues: wrapped, resetOptions, commit };
}
/** Rejects raw RHF methods and cross-suite Root/hook ownership. */
export function requireFormMethodsMetadata(methods: object, suiteIdentity?: symbol): FormMethodsMetadata {
    const metadata = metadataByMethods.get(methods);
    if (!metadata || (suiteIdentity !== undefined && metadata.suiteIdentity !== suiteIdentity))
        throw new TypeError('Form methods must be created by the same Form suite useForm.');
    return metadata;
}
/** Reserves the non-root blocker so RHF cannot mistake validation failure for success. */
export function assertFormFieldName(name: string): void {
    if (name === '__formValidation' || name.startsWith('__formValidation.') || name.startsWith('__formValidation['))
        throw new TypeError('__formValidation is a reserved Form input name.');
}
/** Guards domain record roots at initialization/reset boundaries. */
export function assertFormValues(values: unknown): void {
    if (values && typeof values === 'object' && Object.prototype.hasOwnProperty.call(values, '__formValidation'))
        throw new TypeError('__formValidation is a reserved Form input name.');
}
/** Registers current absolute ownership; re-register on path change while retaining owner identity. */
export function registerFormResetOwner(methods: object, owner: symbol, name: string): () => void {
    const metadata = requireFormMethodsMetadata(methods);
    const registration = Symbol('FormResetRegistration');
    metadata.resetOwners.set(owner, name);
    metadata.mountedResetOwners.set(owner, name);
    metadata.resetOwnerRegistrations.set(owner, registration);
    return () => {
        if (metadata.resetOwnerRegistrations.get(owner) !== registration) return;
        metadata.resetOwnerRegistrations.delete(owner);
        metadata.resetOwners.delete(owner);
        metadata.mountedResetOwners.delete(owner);
        // Path-change cleanup retains queued notifications for this stable owner.
        // The binding's lifetime cleanup removes them on an actual owner unmount.
    };
}
/** Registers a mounted structured array; nested paths move with their enclosing item. */
export function registerFormResetArrayPath(methods: object, name: string): () => void {
    const metadata = requireFormMethodsMetadata(methods);
    const registration = Symbol('FormResetArray');
    metadata.resetArrays.set(registration, name);
    return () => { metadata.resetArrays.delete(registration); };
}
/** Opts a transient query into one reset-scoped handoff across an RHF array remount. */
export function registerFormResetText(methods: object, owner: symbol, name: string, read: () => string | undefined, restore: (text: string) => void): () => void {
    const metadata = requireFormMethodsMetadata(methods);
    metadata.resetTextReaders.set(owner, read);
    const snapshot = metadata.resetTextSnapshots.get(name);
    if (snapshot !== undefined) {
        metadata.resetTextSnapshots.delete(name);
        snapshot.preservedOwners.add(owner);
        restore(snapshot.text);
    }
    return () => { if (metadata.resetTextReaders.get(owner) === read) metadata.resetTextReaders.delete(owner); };
}
/** Remaps only registered identities before RHF changes its array; stores no values or flags. */
export function synchronizeFormResetArrayOperation(methods: object, name: string, operation: FormResetArrayOperation): void {
    const metadata = requireFormMethodsMetadata(methods);
    const removed = operation.type === 'remove' && operation.indices !== undefined ? new Set(operation.indices) : undefined;
    const nextIndex = (index: number): number | undefined => {
        switch (operation.type) {
            case 'append': return index;
            case 'prepend': return index + 1;
            case 'insert': return index >= operation.index ? index + 1 : index;
            case 'remove': return removed === undefined || removed.has(index) ? undefined : index - [...removed].filter(value => value < index).length;
            case 'update': return index === operation.index ? undefined : index;
            case 'replace': return undefined;
            case 'swap': return index === operation.first ? operation.second : index === operation.second ? operation.first : index;
            case 'move':
                if (index === operation.from) return operation.to;
                if (operation.from < operation.to && index > operation.from && index <= operation.to) return index - 1;
                if (operation.from > operation.to && index >= operation.to && index < operation.from) return index + 1;
                return index;
        }
    };
    const remap = (path: string): string | undefined => {
        if (!path.startsWith(`${name}.`)) return path;
        const suffix = path.slice(name.length + 1);
        const match = /^(\d+)(\..*)?$/.exec(suffix);
        if (!match) return path;
        const index = nextIndex(Number(match[1]));
        return index === undefined ? undefined : `${name}.${index}${match[2] ?? ''}`;
    };
    for (const [owner, path] of metadata.resetOwners) {
        const next = remap(path);
        if (next === undefined) {
            metadata.resetOwners.delete(owner);
            metadata.inputValidators.delete(owner);
            metadata.pendingOwnerResets.delete(owner);
            metadata.resetTextReaders.delete(owner);
            metadata.fileOwners.delete(owner);
        } else {
            metadata.resetOwners.set(owner, next);
            const validator = metadata.inputValidators.get(owner);
            if (validator) validator.name = next;
        }
    }
    for (const [registration, path] of metadata.resetArrays) {
        const next = remap(path);
        if (next === undefined) metadata.resetArrays.delete(registration);
        else metadata.resetArrays.set(registration, next);
    }
    const snapshots = [...metadata.resetTextSnapshots];
    metadata.resetTextSnapshots.clear();
    for (const [path, snapshot] of snapshots) {
        const next = remap(path);
        if (next !== undefined) metadata.resetTextSnapshots.set(next, snapshot);
    }
}
/** Delivers targeted notifications once the affected stable owners mount their current paths. */
export function commitFormResetArrayStructure(methods: object, name: string): void {
    const metadata = requireFormMethodsMetadata(methods);
    if (![...metadata.resetArrays.values()].includes(name)) return;
    for (const [owner, events] of metadata.pendingOwnerResets) {
        const current = metadata.resetOwners.get(owner);
        if (current === undefined) { metadata.pendingOwnerResets.delete(owner); continue; }
        if (!current.startsWith(`${name}.`) || metadata.mountedResetOwners.get(owner) !== current) continue;
        metadata.pendingOwnerResets.delete(owner);
        for (const event of events) {
            const notification = { ...event, owners: new Set([owner]) };
            for (const listener of metadata.listeners) listener(notification);
        }
    }
    // Descendant layout effects have consumed valid recipients by this point.
    for (const path of metadata.resetTextSnapshots.keys())
        if (path.startsWith(`${name}.`)) metadata.resetTextSnapshots.delete(path);
}
/** Listens only to this form's reset notifications. */
export function subscribeToFormResets(methods: object, listener: (event: FormResetEvent) => void): () => void {
    const metadata = requireFormMethodsMetadata(methods);
    metadata.listeners.add(listener);
    return () => { metadata.listeners.delete(listener); };
}
/** Emits after RHF has performed the reset; never applies another RHF reset. */
export function notifyFormReset<TInput extends FieldValues, TContext, TOutput>(methods: UseFormReturn<TInput, TContext, TOutput>, options: {
    name?: string;
    preserveValues?: boolean;
    keepDirtyValues?: boolean;
    source?: FormResetEvent['source'];
    preservedOwners?: ReadonlySet<symbol>;
} = {}): void {
    const metadata = requireFormMethodsMetadata(methods);
    const preservedOwners = new Set(options.preservedOwners);
    if (options.keepDirtyValues)
        for (const [owner, name] of metadata.resetOwners)
            if (methods.getFieldState(name as never).isDirty)
                preservedOwners.add(owner);
    const owners = new Set<symbol>();
    const event: FormResetEvent = { version: metadata.resetEvent.version + 1, name: options.name, preserveValues: options.preserveValues === true, preservedOwners, source: options.source ?? 'explicit', owners };
    for (const [owner, name] of metadata.resetOwners) {
        if (!formResetTargets(event, name)) continue;
        if (metadata.mountedResetOwners.get(owner) !== name || metadata.pendingOwnerResets.has(owner)) {
            const pending = metadata.pendingOwnerResets.get(owner) ?? [];
            pending.push(event);
            metadata.pendingOwnerResets.set(owner, pending);
        } else owners.add(owner);
    }
    metadata.resetEvent = event;
    for (const listener of metadata.listeners)
        listener(event);
}
/** Targets a field or any descendant under an object/array reset name. */
export function formResetTargets(event: FormResetEvent, name: string): boolean {
    return event.name === undefined || event.name === name || name.startsWith(`${event.name}.`);
}
/** Enhances public reset/register/value boundaries while preserving RHF's state engine. */
export function useFormMethods<TInput extends FieldValues, TContext, TOutput>(methods: UseFormReturn<TInput, TContext, TOutput>, suiteIdentity: symbol, inputValidators: FormInputValidators, hasResolver: boolean, normalizeRules: typeof normalizeFormRules, getStrings: () => FormStrings): FormMethods<TInput, TContext, TOutput> {
    const metadata = React.useMemo<FormMethodsMetadata>(() => {
        const setValue = methods.setValue;
        return {
            suiteIdentity, getStrings, stringsListeners: new Set(), inputValidators, hasResolver, isResetting: false, valueWriteDepth: 0,
            resetOwners: new Map(), mountedResetOwners: new Map(), resetOwnerRegistrations: new Map(),
            resetArrays: new Map(), pendingOwnerResets: new Map(),
            resetTextReaders: new Map(), resetTextSnapshots: new Map(), listeners: new Set(),
            fileOwners: new Map(),
            // FileInput eligibility establishes its nullable write domain; use the original public
            // setter to avoid recursively invoking package normalization before RHF's mount effect.
            clearFileValue: name => setValue(name as never, clearedFileSelection as never),
            resetEvent: { version: 0, preserveValues: false, preservedOwners: new Set(), source: 'explicit' },
        };
    }, [methods]);
    metadata.hasResolver = hasResolver;
    metadataByMethods.set(methods, metadata);
    const originals = React.useRef({ reset: methods.reset, resetField: methods.resetField, register: methods.register, unregister: methods.unregister, setValue: methods.setValue, setValues: methods.setValues, resetDefaultValues: methods.resetDefaultValues });
    const wrappers = React.useMemo(() => ({
        reset: ((values, options) => {
            if (typeof values !== 'function')
                assertFormValues(values);
            const effectiveOptions = { keepValues: false, keepDirtyValues: false, ...metadata.committedResetOptions, ...options };
            metadata.automaticReset = undefined;
            const preservedOwners = prepareFormReset(methods, effectiveOptions);
            metadata.isResetting = true;
            let didReset = false;
            try {
                originals.current.reset(typeof values === 'function' ? current => { const next = values(current); assertFormValues(next); return next; } : values, effectiveOptions);
                didReset = true;
                synchronizeFormFileOwners(methods, undefined, preservedOwners, effectiveOptions.keepValues);
                notifyFormReset(methods, { preserveValues: effectiveOptions.keepValues, preservedOwners });
            }
            finally {
                if (!didReset) metadata.resetTextSnapshots.clear();
                metadata.isResetting = false;
            }
        }) as typeof methods.reset,
        resetField: ((name, options) => {
            assertFormFieldName(name);
            // A targeted reset supersedes a dirty-query handoff before new RHF row keys mount.
            for (const path of metadata.resetTextSnapshots.keys())
                if (path === name || path.startsWith(`${name}.`)) metadata.resetTextSnapshots.delete(path);
            originals.current.resetField(name, options);
            synchronizeFormFileOwners(methods, name);
            notifyFormReset(methods, { name });
        }) as typeof methods.resetField,
        register: ((name, options) => {
            assertFormFieldName(name);
            if (metadata.hasResolver && options && ['required', 'min', 'max', 'minLength', 'maxLength', 'pattern', 'validate'].some(key => Object.prototype.hasOwnProperty.call(options, key)))
                throw new TypeError('Form field rules cannot accompany a schema/resolver validation owner.');
            return originals.current.register(name, normalizeRules(options as RegisterOptions<TInput> | undefined, undefined, () => metadata.getStrings().unexpectedFailure) as typeof options);
        }) as typeof methods.register,
        setValues: ((values, options) => {
            if (typeof values !== 'function') assertFormValues(values);
            withFormValueWrite(metadata, () => {
                originals.current.setValues(typeof values === 'function' ? current => {
                    const next = values(current);
                    assertFormValues(next);
                    return next;
                } : values, options);
                synchronizeFormFileOwners(methods, undefined, undefined, true);
            });
        }) as typeof methods.setValues,
        resetDefaultValues: ((values, options) => {
            assertFormValues(values);
            withFormValueWrite(metadata, () => originals.current.resetDefaultValues(values, options));
        }) as typeof methods.resetDefaultValues,
        unregister: ((name, options) => withFormValueWrite(metadata, () => originals.current.unregister(name, options))) as typeof methods.unregister,
        setValue: ((name, value, options) => {
            assertFormFieldName(name);
            withFormValueWrite(metadata, () => {
                originals.current.setValue(name, value, options);
                synchronizeFormFileOwners(methods, name, undefined, true);
            });
        }) as typeof methods.setValue,
    }), [methods, metadata]);
    Object.assign(methods, wrappers);
    return methods as FormMethods<TInput, TContext, TOutput>;
}
/** Observes RHF's public reset broadcasts for reactive values/async defaults.
 * The hook never invokes reset itself: RHF remains the sole initialization owner.
 * RHF does not tag reset broadcasts. A committed initialization preparation plus an
 * unnamed notification outside enhanced public value writes acknowledges its reset.
 * Equal reactive records discard unacknowledged preparation after RHF's values effect.
 */
export function useFormResetSynchronization<TInput extends FieldValues, TContext, TOutput>(methods: UseFormReturn<TInput, TContext, TOutput>, options: {
    values?: TInput;
    resetOptions?: Parameters<typeof methods.reset>[1];
}, defaultsBridge?: FormResetDefaultsBridge<TInput, TContext, TOutput>): void {
    const committed = React.useRef(options);
    // RHF's public values effect performs a genuine reset on first initialization too.
    const previousValues = React.useRef<TInput>();
    const reactivePending = React.useRef(false);
    const isLoading = methods.formState.isLoading;
    const asyncPending = React.useRef(isLoading);
    React.useLayoutEffect(() => {
        committed.current = options;
        const metadata = requireFormMethodsMetadata(methods);
        metadata.committedResetOptions = options.resetOptions === undefined ? undefined : { ...options.resetOptions };
        defaultsBridge?.commit(methods, metadata.committedResetOptions);
        reactivePending.current = options.values !== undefined && options.values !== previousValues.current;
        previousValues.current = options.values;
        if (reactivePending.current)
            metadata.automaticReset = { preservedOwners: prepareFormReset(methods, options.resetOptions), preserveValues: options.resetOptions?.keepValues === true, source: 'reactive' };
    });
    React.useLayoutEffect(() => methods.subscribe({ formState: { values: true }, callback: state => {
            const metadata = requireFormMethodsMetadata(methods);
            const preparation = metadata.automaticReset;
            if (preparation && !metadata.isResetting && metadata.valueWriteDepth === 0 && state.name === undefined && state.type === undefined) {
                reactivePending.current = false;
                if (preparation.source === 'async') asyncPending.current = false;
                metadata.automaticReset = undefined;
                synchronizeFormFileOwners(methods, undefined, preparation.preservedOwners, preparation.preserveValues);
                notifyFormReset(methods, preparation);
            }
        } }), [methods]);
    React.useEffect(() => {
        if (asyncPending.current && !isLoading) {
            asyncPending.current = false;
            const metadata = requireFormMethodsMetadata(methods);
            const preparation = metadata.automaticReset;
            metadata.automaticReset = undefined;
            synchronizeFormFileOwners(methods, undefined, preparation?.preservedOwners, preparation?.preserveValues);
            notifyFormReset(methods, preparation ?? { source: 'async', preserveValues: committed.current.resetOptions?.keepValues });
        }
    }, [methods, isLoading]);
    // RHF's values effect runs earlier in this hook stack. If it emitted no reset, discard intent.
    React.useEffect(() => {
        if (reactivePending.current) {
            const metadata = requireFormMethodsMetadata(methods);
            if (metadata.automaticReset?.source === 'reactive') {
                metadata.automaticReset = undefined;
                metadata.resetTextSnapshots.clear();
            }
        }
        reactivePending.current = false;
    });
}
