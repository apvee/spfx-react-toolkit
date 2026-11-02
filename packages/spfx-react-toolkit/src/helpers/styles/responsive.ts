/**
 * Inclusive container inline-size queries at 480, 640 and 1024px. Apply container.inlineSize to an ancestor; an element cannot query itself.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A reusable widget must adapt to its named ancestor’s available inline size, even when the viewport stays unchanged.
 * @example Basic usage
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.small(hover(foreground.subtle));
 * ```
 *
 * @example Resize a named container at each breakpoint — inclusive small/medium/large with a real ancestor
 * ```tsx
 * import * as React from 'react';
 * import { Field, Input, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, flex, gap, padding, presets, responsive, container } from '@apvee/spfx-react-toolkit';
 * interface DemoRegions {
 *   controls?: React.ReactNode;
 *   commands?: React.ReactNode;
 *   status?: React.ReactNode;
 *   children?: React.ReactNode;
 *   diagnostics?: React.ReactNode;
 * }
 * function standalone(regions: DemoRegions): React.ReactElement {
 *   return <div>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</div>;
 * }
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [threshold, setThreshold] = React.useState<'small' | 'medium' | 'large'>('medium');
 *   const [size, setSize] = React.useState(420);
 *   const [draft, setDraft] = React.useState('420');
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setThreshold('medium'); setSize(420); setDraft('420');
 *   }, [resetRevision]);
 *   return render({
 *     controls: <>
 *       <Field label="Breakpoint"><Select value={threshold} onChange={(_, data) => {
 *         if (data.value === 'small' || data.value === 'medium' || data.value === 'large') setThreshold(data.value);
 *       }}><option value="small">small</option><option value="medium">medium</option><option value="large">large</option></Select></Field>
 *       <Field label="Ancestor width"><Input type="number" min={0} step="any" value={draft} onChange={(_, data) => {
 *         setDraft(data.value);
 *         if (data.value.trim() && Number.isFinite(Number(data.value)) && Number(data.value) >= 0) setSize(Number(data.value));
 *       }} onBlur={() => setDraft(String(size))} /></Field>
 *     </>,
 *     status: <p>Inclusive thresholds: small 480px, medium 640px, large 1024px. The child measures its nearest named ancestor, never itself.</p>,
 *     children: <section role="region" aria-label="responsive query preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
 *       <div data-style-helper="responsive-ancestor" className={sx(container.inlineSize, width.px(size))}>
 *         <div data-style-helper="responsive" className={sx(width.px(180), flex.column, gap.small, padding.small, presets.alternative,
 *           responsive[threshold](width.px(300), flex.row))}><span>First item</span><span>Second item</span></div>
 *       </div>
 *     </section>,
 *     diagnostics: <p>Reset restores medium and the 420px ancestor without replacing the Demo or host renderer.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select small, medium or large and compare threshold minus one with its inclusive limit. Change ancestor width while the viewport remains fixed. Reset restores the displayed baseline.',
 * };
 * ```
 */
import { SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { createResponsive } from './scopes.internal';

/**
 * Applies descriptors at the nearest ancestor named apvee-sx at inline-size \>= 480px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.small(hover(foreground.subtle));
 * ```
 */
export function small(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 480, inputs);
}

/**
 * Applies descriptors at the nearest ancestor named apvee-sx at inline-size \>= 640px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.medium(hover(foreground.subtle));
 * ```
 */
export function medium(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 640, inputs);
}

/**
 * Applies descriptors at the nearest ancestor named apvee-sx at inline-size \>= 1024px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable container query descriptor.
 * @example
 * ```ts
 * import { responsive, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = responsive.large(hover(foreground.subtle));
 * ```
 */
export function large(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('container', 1024, inputs);
}
