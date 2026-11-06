/**
 * Individual flex/grid item alignment declarations; these do not establish layout.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen One child needs a cross-axis alignment different from its flex/grid container’s other items.
 * @example Basic usage
 * ```ts
 * import { alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = alignSelf.auto;
 * ```
 *
 * @scenario Select an individual flex item alignment — subject is a child, not the parent
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, flex, gap, padding, presets, alignSelf, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = alignSelf;
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
 *     status: <p aria-live="polite">Applied: alignSelf.{member}.</p>,
 *     children: <section role="region" aria-label="alignSelf descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div className={sx(flex.row, width.px(240), height.px(150), gap.small)}><span className={sx(width.px(140), presets.canvas)}>Sibling</span><div data-style-helper="alignSelf" className={sx(width.px(140), padding.small, presets.alternative, descriptor)}>Actual alignSelf subject</div><span className={sx(width.px(140), presets.canvas)}>Sibling</span></div></section>,
 *     diagnostics: <p>Select an individual flex item alignment; subject is a child, not the parent. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select an individual flex item alignment — subject is a child, not the parent. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets alignSelf to auto without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'auto', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to start without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.start)}>Content</div>;
 * }
 * ```
 */
export const start: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'start', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to center without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.center)}>Content</div>;
 * }
 * ```
 */
export const center: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'center', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to end without changing display. Uses logical CSS alignment, respecting writing mode and direction.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.end)}>Content</div>;
 * }
 * ```
 */
export const end: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'end', { property: 'alignSelf', fallback: 'auto' });

/** Sets alignSelf to stretch without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, alignSelf } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(alignSelf.stretch)}>Content</div>;
 * }
 * ```
 */
export const stretch: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('alignSelf', 'stretch', { property: 'alignSelf', fallback: 'auto' });
