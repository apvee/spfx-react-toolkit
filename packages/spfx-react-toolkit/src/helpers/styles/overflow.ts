/**
 * Overflow declarations for both axes; provide constrained dimensions when scrolling is required.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A bounded region must scroll or clip excess content on both axes; set its dimensions separately.
 * @example Basic usage
 * ```ts
 * import { overflow } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = overflow.visible;
 * ```
 *
 * @example Select overflow behavior on both axes — bounded wide/tall content; include axis choices as real public nested members
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, padding, borderWidth, borderStyle, borderColor, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = { visible: overflow.visible, hidden: overflow.hidden, auto: overflow.auto, 'horizontal.visible': overflow.horizontal.visible, 'horizontal.hidden': overflow.horizontal.hidden, 'horizontal.auto': overflow.horizontal.auto, 'vertical.visible': overflow.vertical.visible, 'vertical.hidden': overflow.vertical.hidden, 'vertical.auto': overflow.vertical.auto };
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
 *     status: <p aria-live="polite">Applied: overflow.{member}.</p>,
 *     children: <section role="region" aria-label="overflow descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="overflow" className={sx(width.px(180), height.px(100), borderWidth.thin, borderStyle.solid, borderColor.primary, descriptor)}><div className={sx(width.px(480), height.px(240))}>Wide and tall actual overflow content</div></div></section>,
 *     diagnostics: <p>Select overflow behavior on both axes; bounded wide/tall content; include axis choices as real public nested members. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select overflow behavior on both axes — bounded wide/tall content; include axis choices as real public nested members. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Controls only the horizontal overflow axis. */
export * as horizontal from './overflow-horizontal';
/** Controls only the vertical overflow axis. */
export * as vertical from './overflow-vertical';

/** Sets both overflow axes to visible; adds no dimensions or scrollbar appearance.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.visible)}>Content</div>;
 * }
 * ```
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.visible', [
  /*#__PURE__*/ createDeclaration('overflowX', 'visible', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'visible', { property: 'overflowY', fallback: 'visible' }),
]);

/** Sets both overflow axes to hidden; adds no dimensions or scrollbar appearance.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.hidden)}>Content</div>;
 * }
 * ```
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.hidden', [
  /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' }),
]);

/** Sets both overflow axes to auto; adds no dimensions or scrollbar appearance.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createRecipe('overflow.auto', [
  /*#__PURE__*/ createDeclaration('overflowX', 'auto', { property: 'overflowX', fallback: 'visible' }),
  /*#__PURE__*/ createDeclaration('overflowY', 'auto', { property: 'overflowY', fallback: 'visible' }),
]);
