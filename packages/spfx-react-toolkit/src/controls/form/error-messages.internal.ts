import * as React from 'react';
function isRecord(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== undefined && Object.prototype.toString.call(value) === '[object Object]'; }
function append(messages: string[], value: unknown): void {
    for (const candidate of Array.isArray(value) ? value : [value]) {
        if (typeof candidate === 'string' && candidate !== '' && !messages.includes(candidate))
            messages.push(candidate);
    }
}
/** Reads only message/types, never arbitrary child trees or native ref objects. */
export function getFormErrorMessages(...errors: readonly unknown[]): readonly string[] {
    const messages: string[] = [];
    for (const error of errors) {
        if (!isRecord(error))
            continue;
        append(messages, error.message);
        if (isRecord(error.types))
            for (const key of Object.keys(error.types).sort())
                append(messages, error.types[key]);
    }
    return messages;
}
/** Collects only own root metadata and one level of named global child errors. */
export function getGlobalFormErrorMessages(root: unknown, blocker: unknown): readonly string[] {
    const children = isRecord(root) ? Object.keys(root).sort().filter(key => !['message', 'type', 'types', 'ref'].includes(key)).map(key => root[key]) : [];
    return getFormErrorMessages(root, ...children, blocker);
}
/** A single text node or a native list when multiple errors are active. */
export function renderFormErrorMessages(messages: readonly string[]): React.ReactNode {
    return messages.length <= 1 ? messages[0] : React.createElement('ul', undefined, messages.map(message => React.createElement('li', { key: message }, message)));
}
