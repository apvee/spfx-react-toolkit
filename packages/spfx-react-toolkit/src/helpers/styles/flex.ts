/**
 * Flex display/direction recipes and wrapping declarations. wrap/noWrap do not set display.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Children need a flex row/column or an existing flex layout needs an explicit wrapping policy.
 * @example Basic usage
 * ```ts
 * import { flex } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = flex.row;
 * ```
 *
 * @scenario Select flex direction and wrapping — actual sized item row/column/wrap
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, flex, gap, padding, presets, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = flex;
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
 *     status: <p aria-live="polite">Applied: flex.{member}.</p>,
 *     children: <section role="region" aria-label="flex descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="flex" className={sx(flex.row, width.px(240), height.px(140), gap.small, presets.alternative, descriptor)}><span className={sx(width.px(100))}>First child</span><span className={sx(width.px(100), height.px(70))}>Second taller child</span><span className={sx(width.px(100))}>Third child</span></div></section>,
 *     diagnostics: <p>Select flex direction and wrapping; actual sized item row/column/wrap. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select flex direction and wrapping — actual sized item row/column/wrap. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Flex container arranged in a row; gap and alignment remain independent.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex, gap } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.row, gap.medium)}>Content</div>;
 * }
 * ```
 */
export const row: SxBaseDescriptor = /*#__PURE__*/ createRecipe('flex.row', [
  /*#__PURE__*/ createDeclaration('display', 'flex', { property: 'display', fallback: 'inline' }),
  /*#__PURE__*/ createDeclaration('flexDirection', 'row', { property: 'flexDirection', fallback: 'row' })
]);

/** Flex container arranged in a column; gap and alignment remain independent.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex, gap } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.column, gap.medium)}>Content</div>;
 * }
 * ```
 */
export const column: SxBaseDescriptor = /*#__PURE__*/ createRecipe('flex.column', [
  /*#__PURE__*/ createDeclaration('display', 'flex', { property: 'display', fallback: 'inline' }),
  /*#__PURE__*/ createDeclaration('flexDirection', 'column', { property: 'flexDirection', fallback: 'row' })
]);

/** Allows flex items to wrap without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.wrap)}>Content</div>;
 * }
 * ```
 */
export const wrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexWrap', 'wrap', { property: 'flexWrap', fallback: 'nowrap' });

/** Keeps flex items on one line without changing display.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flex } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flex.noWrap)}>Content</div>;
 * }
 * ```
 */
export const noWrap: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexWrap', 'nowrap', { property: 'flexWrap', fallback: 'nowrap' });
