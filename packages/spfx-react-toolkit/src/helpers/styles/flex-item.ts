/**
 * Flex grow/shrink declarations for items in a flex container.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An item in an existing flex container must share spare space or control shrinking under pressure.
 * @example Basic usage
 * ```ts
 * import { flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = flexItem.grow;
 * ```
 *
 * @scenario Select flex item growth and shrink — target item among constrained siblings
 * ```tsx
 * import * as React from 'react';
 * import { Checkbox, Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, flex, gap, padding, presets, flexItem, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = flexItem;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [wide, setWide] = React.useState(false);
 *   const [member, setMember] = React.useState(Object.keys(choices)[0]);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setWide(false); setMember(Object.keys(choices)[0]);
 *   }, [resetRevision]);
 *   const descriptor = choices[member];
 *   return render({
 *     controls: <><Checkbox label="Extra row space" checked={wide} onChange={(_, data) => setWide(data.checked === true)} /><Field label="Descriptor"><Select value={member} onChange={(_, data) => {
 *       if (Object.prototype.hasOwnProperty.call(choices, data.value)) setMember(data.value);
 *     }}>{Object.keys(choices).map(key => <option key={key} value={key}>{key}</option>)}</Select></Field>
 *     </>,
 *     status: <p aria-live="polite">Applied: flexItem.{member}.</p>,
 *     children: <section role="region" aria-label="flexItem descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div className={sx(flex.row, width.px(wide ? 640 : 240), height.px(150), gap.small)}><span className={sx(width.px(140), presets.canvas)}>Sibling</span><div data-style-helper="flexItem" className={sx(width.px(140), padding.small, presets.alternative, descriptor)}>Actual flexItem subject</div><span className={sx(width.px(140), presets.canvas)}>Sibling</span></div></section>,
 *     diagnostics: <p>Select flex item growth and shrink; target item among constrained siblings. Toggle extra row space for growth versus a constrained row for shrink. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select flex item growth and shrink — target item among constrained siblings. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets flexGrow to 1 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.grow)}>Content</div>;
 * }
 * ```
 */
export const grow: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexGrow', '1', { property: 'flexGrow', fallback: '0' });

/** Sets flexGrow to 0 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.noGrow)}>Content</div>;
 * }
 * ```
 */
export const noGrow: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexGrow', '0', { property: 'flexGrow', fallback: '0' });

/** Sets flexShrink to 1 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.shrink)}>Content</div>;
 * }
 * ```
 */
export const shrink: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexShrink', '1', { property: 'flexShrink', fallback: '1' });

/** Sets flexShrink to 0 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.noShrink)}>Content</div>;
 * }
 * ```
 */
export const noShrink: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexShrink', '0', { property: 'flexShrink', fallback: '1' });
