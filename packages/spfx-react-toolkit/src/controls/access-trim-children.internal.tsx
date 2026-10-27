import * as React from 'react';

type BranchKind = 'loading' | 'fallback' | 'error';
interface BranchMetadata { readonly family: object; readonly kind: BranchKind; }
const branchMarker = Symbol('access-trim-branch');
type MarkedBranch = React.ComponentType<{ children?: React.ReactNode }> & { readonly [branchMarker]: BranchMetadata };

// Local component metadata only; no global registry, rendering or network effects.
export function createAccessTrimBranch(family: object, kind: BranchKind): MarkedBranch {
  function Branch(props: { children?: React.ReactNode }): React.ReactElement {
    return <>{props.children}</>;
  }
  return Object.assign(Branch, { [branchMarker]: { family, kind } });
}
interface TrimChildren {
  readonly granted: React.ReactNode;
  readonly loading: React.ReactNode;
  readonly fallback: React.ReactNode;
  readonly errorFallback: React.ReactNode;
  readonly branchKeys: Readonly<Partial<Record<BranchKind, React.Key>>>;
  readonly error?: Error;
}
/** Preserve original elements, nested-array structure and keys; never execute consumers. */
export function readAccessTrimChildren(children: React.ReactNode, family: object): TrimChildren {
  const branches: Partial<Record<BranchKind, React.ReactNode>> = {};
  const branchKeys: Partial<Record<BranchKind, React.Key>> = {};
  const seen = new Set<BranchKind>();
  let error: Error | undefined;
  function fragmentContainsBranch(child: React.ReactNode): boolean {
    if (Array.isArray(child)) return child.some(fragmentContainsBranch);
    if (!React.isValidElement<{ children?: React.ReactNode }>(child)) return false;
    if (typeof child.type === 'function' && (child.type as Partial<MarkedBranch>)[branchMarker]) return true;
    return child.type === React.Fragment && fragmentContainsBranch(child.props.children);
  }
  function visit(child: React.ReactNode): React.ReactNode {
    if (Array.isArray(child)) return child.map(visit);
    if (!React.isValidElement<{ children?: React.ReactNode }>(child)) return child;
    const type = child.type;
    if (type === React.Fragment && fragmentContainsBranch(child.props.children)) {
      error = new Error('Access trim branches must be direct children, without a fragment wrapper.');
      return undefined;
    }
    const metadata = typeof type === 'function'
      ? (type as Partial<MarkedBranch>)[branchMarker]
      : undefined;
    if (!metadata) return child;
    if (metadata.family !== family) {
      error = new Error('Access trim branches must belong to their containing control.');
    } else if (seen.has(metadata.kind)) {
      error = new Error('Access trim branches must be unique direct children.');
    } else {
      seen.add(metadata.kind);
      branches[metadata.kind] = child.props.children;
      // ReactElement.key uses legacy null for an omitted consumer key.
      if (child.key !== null) branchKeys[metadata.kind] = child.key;
    }
    return undefined;
  }
  const granted = visit(children);
  return { granted, loading: branches.loading, fallback: branches.fallback, errorFallback: branches.error, branchKeys, error };
}
export function renderAccessTrim(children: TrimChildren, state: { readonly granted: boolean; readonly loading: boolean; readonly error?: Error }): React.ReactElement {
  const role = state.error || children.error ? 'error'
    : state.loading ? 'loading' : state.granted ? 'granted' : 'fallback';
  const content = role === 'error' ? children.errorFallback : role === 'loading' ? children.loading
    : role === 'granted' ? children.granted : children.fallback;
  // Isolate branch lifetimes without adding DOM. The same unkeyed child type in
  // another branch must mount anew; an unchanged role/marker key preserves state.
  const key = role === 'granted' ? role : `${role}:${JSON.stringify(children.branchKeys[role])}`;
  return <React.Fragment key={key}>{content}</React.Fragment>;
}
export function validateTrimMatch(value: unknown): 'all' | 'any' {
  if (value === undefined || value === 'all') return 'all';
  if (value === 'any') return 'any';
  throw new Error('Access trim match must be all or any.');
}
