/**
 * Inclusive viewport width queries at 480, 640 and 1024px, independent of container size.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Page-level layout must follow window width independently of the widths of nested cards or sidebars.
 * @example Basic usage
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.small(hover(foreground.subtle));
 * ```
 *
 * @example Resize the viewport at each breakpoint — inclusive small/medium/large
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, flex, gap, padding, presets, viewport } from '@apvee/spfx-react-toolkit';
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
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setThreshold('medium');
 *   }, [resetRevision]);
 *   return render({
 *     controls: <>
 *       <Field label="Breakpoint"><Select value={threshold} onChange={(_, data) => {
 *         if (data.value === 'small' || data.value === 'medium' || data.value === 'large') setThreshold(data.value);
 *       }}><option value="small">small</option><option value="medium">medium</option><option value="large">large</option></Select></Field>
 *     </>,
 *     status: <p>Inclusive thresholds: small 480px, medium 640px, large 1024px. Resize the actual viewport; this query does not measure the ancestor.</p>,
 *     children: <section role="region" aria-label="viewport query preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
 *       <div data-style-helper="viewport-ancestor" className={sx(width.px(420))}>
 *         <div data-style-helper="viewport" className={sx(width.px(180), flex.column, gap.small, padding.small, presets.alternative,
 *           viewport[threshold](width.px(300), flex.row))}><span>First item</span><span>Second item</span></div>
 *       </div>
 *     </section>,
 *     diagnostics: <p>Reset restores medium without replacing the Demo or host renderer.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select small, medium or large and compare threshold minus one with its inclusive limit. Resize the actual viewport independently of the fixed ancestor. Reset restores the displayed baseline.',
 * };
 * ```
 */
import { SxBaseDescriptor, SxStateDescriptor, SxResponsiveDescriptor } from './types';
import { createResponsive } from './scopes.internal';

/**
 * Applies descriptors at viewport width \>= 480px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.small(hover(foreground.subtle));
 * ```
 */
export function small(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 480, inputs);
}

/**
 * Applies descriptors at viewport width \>= 640px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.medium(hover(foreground.subtle));
 * ```
 */
export function medium(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 640, inputs);
}

/**
 * Applies descriptors at viewport width \>= 1024px.
 * @param inputs - Base or state descriptors; nested queries and class strings are unsupported.
 * @returns Immutable viewport query descriptor.
 * @example
 * ```ts
 * import { viewport, foreground, hover } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = viewport.large(hover(foreground.subtle));
 * ```
 */
export function large(...inputs: readonly (SxBaseDescriptor | SxStateDescriptor)[]): SxResponsiveDescriptor {
  return createResponsive('viewport', 1024, inputs);
}
