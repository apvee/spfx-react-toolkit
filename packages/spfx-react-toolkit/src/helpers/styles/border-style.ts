/**
 * Four-side border styles; compose width and color as needed.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A box needs a solid border style on all sides while width and color remain separately selectable.
 * @example Basic usage
 * ```ts
 * import { borderStyle } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = borderStyle.solid;
 * ```
 *
 * @example Select border style descriptors — keep explicit width/color
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, height, minWidth, overflow, padding, presets, borderWidth, borderStyle, borderColor, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = borderStyle;
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
 *     status: <p aria-live="polite">Applied: borderStyle.{member}.</p>,
 *     children: <section role="region" aria-label="borderStyle descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="borderStyle" className={sx(width.px(180), height.px(100), padding.small, borderWidth.thick, borderStyle.solid, borderColor.primary, presets.canvas, descriptor)}>Actual borderStyle subject</div></section>,
 *     diagnostics: <p>Select border style descriptors; keep explicit width/color. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select border style descriptors — keep explicit width/color. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies solid to all four border sides. Compose border width, style, and color explicitly.
 * Token values require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, borderStyle } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(borderStyle.solid)}>Content</div>;
 * }
 * ```
 */
export const solid: SxBaseDescriptor = /*#__PURE__*/ createRecipe('borderStyle.solid', [
  /*#__PURE__*/ createDeclaration('borderTopStyle', 'solid', { property: 'borderTopStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderRightStyle', 'solid', { property: 'borderRightStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderBottomStyle', 'solid', { property: 'borderBottomStyle', fallback: 'none' }),
  /*#__PURE__*/ createDeclaration('borderLeftStyle', 'solid', { property: 'borderLeftStyle', fallback: 'none' }),
]);
