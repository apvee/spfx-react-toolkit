/**
 * Fluent foreground color descriptors. Color alone adds no link behavior or disabled semantics.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Text or icons need a Fluent foreground role without also changing their surface background.
 * @example Basic usage
 * ```ts
 * import { foreground } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = foreground.primary;
 * ```
 *
 * @example Select foreground role descriptors — actual text and disabled aria
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, foreground, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = foreground;
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
 *     status: <p aria-live="polite">Applied: foreground.{member}.</p>,
 *     children: <section role="region" aria-label="foreground descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}><p data-style-helper="foreground" aria-disabled={member === 'disabled'} className={sx(padding.medium, descriptor)}>Actual foreground subject</p></section>,
 *     diagnostics: <p>Select foreground role descriptors; actual text and disabled aria. Token descriptors inherit the host theme; Reset restores the first choice without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select foreground role descriptors — actual text and disabled aria. Select actual namespace members and compare their effect on the actual subject. Reset restores the first choice while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies Fluent's colorNeutralForeground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.primary)}>Content</div>;
 * }
 * ```
 */
export const primary: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground1, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForeground2 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.subtle)}>Content</div>;
 * }
 * ```
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground2, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForeground3 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.muted)}>Content</div>;
 * }
 * ```
 */
export const muted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground3, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundDisabled CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.disabled)}>Content</div>;
 * }
 * ```
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundDisabled, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorBrandForeground1 CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.brand)}>Content</div>;
 * }
 * ```
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorBrandForeground1, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundOnBrand CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.onBrand)}>Content</div>;
 * }
 * ```
 */
export const onBrand: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundOnBrand, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorNeutralForegroundInverted CSS variable.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.inverted)}>Content</div>;
 * }
 * ```
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundInverted, { property: 'color', fallback: 'inherit' });

/** Applies Fluent's colorBrandForegroundLink CSS variable. Color only; adds no link semantics.
 * Requires theme variables in scope; adds no state or geometry.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, foreground } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(foreground.link)}>Content</div>;
 * }
 * ```
 */
export const link: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('color', tokens.colorBrandForegroundLink, { property: 'color', fallback: 'inherit' });
