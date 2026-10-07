import { getDescriptorData, SxBinding, SxBreakpoint, SxProperty, SxState } from './descriptor.internal';
import { SxDescriptor, SxInput } from './types';

export type SxQuery = { readonly target: 'base' } | {
  readonly target: 'container' | 'viewport';
  readonly breakpoint: SxBreakpoint;
};
export interface SxNormalizedDeclaration {
  readonly property: SxProperty;
  readonly value: string;
  readonly query: SxQuery;
  readonly state?: SxState;
  readonly binding: SxBinding;
}
export type SxNormalizedSegment =
  | { readonly kind: 'class'; readonly className: string }
  | { readonly kind: 'declarations'; readonly declarations: readonly SxNormalizedDeclaration[] };

/** Stable scope identity; values never participate in property conflict detection. */
export function sxScopeKey(query: SxQuery, state?: SxState): string {
  const queryKey = query.target === 'base' ? 'base' : `${query.target}-${query.breakpoint}`;
  return state ? `${queryKey}-${state}` : queryKey;
}
/** Expand recipes in place and deduplicate only inside contiguous descriptor segments. */
export function normalizeSxInputs(inputs: readonly SxInput[]): readonly SxNormalizedSegment[] {
  const segments: SxNormalizedSegment[] = [];
  let declarations = new Map<string, SxNormalizedDeclaration>();
  const flush = (): void => {
    if (declarations.size) {
      segments.push({ kind: 'declarations', declarations: [...declarations.values()] });
      declarations = new Map();
    }
  };
  const visit = (descriptor: SxDescriptor, query: SxQuery, state?: SxState): void => {
    const data = getDescriptorData(descriptor);
    switch (data.kind) {
      case 'declaration':
        declarations.set(`${data.property}/${sxScopeKey(query, state)}`, {
          property: data.property, value: data.value, query, state, binding: data.binding
        });
        break;
      case 'recipe':
        data.declarations.forEach(input => visit(input, query, state));
        break;
      case 'state':
        data.inputs.forEach(input => visit(input, query, data.state));
        break;
      case 'responsive':
        data.inputs.forEach(input => visit(input, { target: data.target, breakpoint: data.breakpoint }));
        break;
    }
  };
  for (const input of inputs) {
    if (!input) continue;
    if (typeof input === 'string') {
      flush();
      segments.push({ kind: 'class', className: input });
    } else {
      visit(input, { target: 'base' });
    }
  }
  flush();
  return segments;
}
