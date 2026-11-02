/**
 * Fluent background color descriptors. These add no geometry, foreground color or interaction states.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface needs a Fluent background role while retaining its existing foreground and geometry.
 * @example Basic usage
 * ```ts
 * import { background } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = background.canvas;
 * ```
 *
 * @example Select background role descriptors — actual surface and disabled aria
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, background, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = background;
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
 *     status: <p aria-live="polite">Applied: background.{member}.</p>,
 *     children: <section role="region" aria-label="background descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><div data-style-helper="background" aria-disabled={member === 'disabled'} className={sx(padding.medium, descriptor)}>Actual background subject</div></section>,
 *     diagnostics: <p>Select background role descriptors; actual surface and disabled aria. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select background role descriptors — actual surface and disabled aria. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies Fluent's colorNeutralBackground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.canvas)}>Content</div>;
 * }
 * ```
 */
export const canvas: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackground1, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.alternative)}>Content</div>;
 * }
 * ```
 */
export const alternative: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackground2, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorSubtleBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.subtle)}>Content</div>;
 * }
 * ```
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorSubtleBackground, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorTransparentBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.transparent)}>Content</div>;
 * }
 * ```
 */
export const transparent: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorTransparentBackground, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorBrandBackground CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.brand)}>Content</div>;
 * }
 * ```
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorBrandBackground, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorBrandBackground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.brandTint)}>Content</div>;
 * }
 * ```
 */
export const brandTint: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorBrandBackground2, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackgroundInverted CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.inverted)}>Content</div>;
 * }
 * ```
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackgroundInverted, { property: 'backgroundColor', fallback: 'transparent' });

/** Applies Fluent's colorNeutralBackgroundDisabled CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, background } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(background.disabled)}>Content</div>;
 * }
 * ```
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackgroundDisabled, { property: 'backgroundColor', fallback: 'transparent' });
