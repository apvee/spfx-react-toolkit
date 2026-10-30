import * as React from 'react';
const FieldPathContext = React.createContext('');
const DisabledScopeContext = React.createContext(false);
/** Joins an absolute prefix and relative path. */
export function joinFieldPath(parent: string, child: string): string {
    return parent === '' ? child : child === '' ? parent : `${parent}.${child}`;
}
/** Resolves a name in its nearest object/array scope. */
export function useResolvedFieldPath(name: string): string {
    return joinFieldPath(React.useContext(FieldPathContext), name);
}
/** Reads additive disabled state from private composition scopes. */
export function useFieldScopeDisabled(): boolean { return React.useContext(DisabledScopeContext); }
/** Provides an already absolute path prefix; disabled scopes can only add restrictions. */
export function FieldPathScope({ prefix, children, disabled }: {
    prefix: string;
    children?: React.ReactNode;
    disabled?: boolean;
}): React.ReactElement {
    const parentDisabled = useFieldScopeDisabled();
    return <FieldPathContext.Provider value={prefix}><DisabledScopeContext.Provider value={parentDisabled || disabled === true}>{children}</DisabledScopeContext.Provider></FieldPathContext.Provider>;
}
