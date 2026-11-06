/**
 * Main-axis alignment declarations; logical start/end follow direction and writing mode.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An existing flex/grid container needs to position or distribute children along its main axis.
 * @example Basic usage
 * ```ts
 * import { justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = justifyContent.start;
 * ```
 *
 * @scenario Select main axis item alignment — spare inline space and real siblings
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, flex, gap, padding, presets, justifyContent, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = justifyContent;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember(Object.keys(choices)[0]);
 *   }, [resetRevision]);
 *   const descriptor = choices[member];
 *   return render({
 *     controls: <><Field label="Descriptor"><Select value={member} onChange={(_, data) => {
 *       if (Object.prototype.hasOwnProperty.call(choices, data.value)) setMember(data.value);
 *     }}>{Object.keys(choices).map(key => <option key={key} value={key}>{key}</option>)}</Select></Field>
 *     </>,
 *     status: <p aria-live="polite">Applied: justifyContent.{member}.</p>,
 *     children: <section role="region" aria-label="justifyContent descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="justifyContent" className={sx(flex.row, width.px(480), height.px(140), gap.small, presets.alternative, descriptor)}><span className={sx(width.px(100))}>First child</span><span className={sx(width.px(100), height.px(70))}>Second taller child</span><span className={sx(width.px(100))}>Third child</span></div></section>,
 *     diagnostics: <p>Select main axis item alignment; spare inline space and real siblings. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select main axis item alignment — spare inline space and real siblings. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets justifyContent to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'start', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to center without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'center', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'end', { property: 'justifyContent', fallback: 'normal' });

/** Sets justifyContent to space-between without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, justifyContent } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(justifyContent.spaceBetween)}>Content</div>;
 * }
 * ```
 */
export const spaceBetween: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('justifyContent', 'space-between', { property: 'justifyContent', fallback: 'normal' });
