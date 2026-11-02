/**
 * Fluent shadow recipes; these do not add surface geometry or colors.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface needs Fluent elevation styling while its background, dimensions and padding remain separately controlled.
 * @example Basic usage
 * ```ts
 * import { boxShadow } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = boxShadow.none;
 * ```
 *
 * @example Select shadow elevation descriptors — actual surface with room for shadow
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, padding, presets, boxShadow, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = boxShadow;
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
 *     status: <p aria-live="polite">Applied: boxShadow.{member}.</p>,
 *     children: <section role="region" aria-label="boxShadow descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="boxShadow" className={sx(width.px(180), height.px(100), padding.medium, presets.canvas, descriptor)}>Actual boxShadow subject</div></section>,
 *     diagnostics: <p>Select shadow elevation descriptors; actual surface with room for shadow. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select shadow elevation descriptors — actual surface with room for shadow. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies box-shadow none without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, boxShadow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(boxShadow.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', 'none', { property: 'boxShadow', fallback: 'none' });

/** Applies box-shadow var(--shadow4) without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, boxShadow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(boxShadow.small)}>Content</div>;
 * }
 * ```
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', tokens.shadow4, { property: 'boxShadow', fallback: 'none' });

/** Applies box-shadow var(--shadow8) without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, boxShadow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(boxShadow.medium)}>Content</div>;
 * }
 * ```
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', tokens.shadow8, { property: 'boxShadow', fallback: 'none' });
