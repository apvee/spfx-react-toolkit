/**
 * Named inline-size containment for ancestor queries. Adds no dimensions or display; provide host geometry because containment affects intrinsic sizing.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An ancestor must supply the named size reference for responsive descendants; provide its layout geometry separately.
 * @example Basic usage
 * ```ts
 * import { container } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = container.inlineSize;
 * ```
 *
 * @scenario Toggle a named inline size query ancestor — nested child visibly responds when ancestor establishment changes
 * ```tsx
 * import * as React from 'react';
 * import { Checkbox, Field, Input, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, flex, gap, padding, presets, container, responsive } from '@apvee/spfx-react-toolkit';
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
 *   const [established, setEstablished] = React.useState(true);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setThreshold('medium'); setSize(420); setDraft('420'); setEstablished(true);
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
 *       <Checkbox label="Establish named container" checked={established} onChange={(_, data) => setEstablished(data.checked === true)} />
 *     </>,
 *     status: <p>Inclusive thresholds: small 480px, medium 640px, large 1024px. The child measures its nearest named ancestor, never itself.</p>,
 *     children: <section role="region" aria-label="container query preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
 *       <div data-style-helper="container-ancestor" className={sx(established && container.inlineSize, width.px(size))}>
 *         <div data-style-helper="container" className={sx(width.px(180), flex.column, gap.small, padding.small, presets.alternative,
 *           responsive[threshold](width.px(300), flex.row))}><span>First item</span><span>Second item</span></div>
 *       </div>
 *     </section>,
 *     diagnostics: <p>Reset restores medium and the 420px ancestor without replacing the Demo or host renderer.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select small, medium or large and compare threshold minus one with its inclusive limit. Toggle named containment at 720px to switch the child between its 180px column and 300px row. Reset restores the displayed baseline.',
 * };
 * ```
 */
import type { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/**
 * Establishes the named `apvee-sx` inline-size query container for descendants.
 * Adds no width, height, display or spacing. Responsive descriptors measure the
 * nearest named ancestor's inline size, with inclusive 480/640/1024px thresholds.
 * An element never queries its own size; vertical writing modes measure height.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, container, responsive, flex } from '@apvee/spfx-react-toolkit/styles';
 * export function ResponsiveRows(): React.ReactElement {
 *   const sx = useSx();
 *   return <section className={sx(container.inlineSize)}>
 *     <div className={sx(flex.column, responsive.medium(flex.row))}>Content</div>
 *   </section>;
 * }
 * ```
 */
export const inlineSize: SxBaseDescriptor = /*#__PURE__*/ createRecipe('container.inlineSize', [
  /*#__PURE__*/ createDeclaration('containerType', 'inline-size', { property: 'containerType', fallback: 'normal' }),
  /*#__PURE__*/ createDeclaration('containerName', 'apvee-sx', { property: 'containerName', fallback: 'none' })
]);
