/**
 * Cross-axis alignment declarations; these do not establish a flex or grid layout.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An existing flex/grid container needs a common cross-axis alignment for its children.
 * @example Basic usage
 * ```ts
 * import { alignItems } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = alignItems.start;
 * ```
 *
 * @example Select cross axis item alignment — actual fixed-height flex parent
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, flex, gap, padding, presets, alignItems, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = alignItems;
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
 *     status: <p aria-live="polite">Applied: alignItems.{member}.</p>,
 *     children: <section role="region" aria-label="alignItems descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="alignItems" className={sx(flex.row, width.px(320), height.px(140), gap.small, presets.alternative, descriptor)}><span className={sx(width.px(100))}>First child</span><span className={sx(width.px(100), height.px(70))}>Second taller child</span><span className={sx(width.px(100))}>Third child</span></div></section>,
 *     diagnostics: <p>Select cross axis item alignment; actual fixed-height flex parent. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select cross axis item alignment — actual fixed-height flex parent. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets alignItems to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'start', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to center without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'center', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'end', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to stretch without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.stretch)}>Content</div>;
 * }
 * ```
 */
export const stretch: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'stretch', { property: 'alignItems', fallback: 'normal' });

/** Sets alignItems to baseline without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignItems } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignItems.baseline)}>Content</div>;
 * }
 * ```
 */
export const baseline: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignItems', 'baseline', { property: 'alignItems', fallback: 'normal' });
