/**
 * Paired Fluent background/foreground recipes; transparent leaves foreground inherited. Recipes add no geometry or interaction variants. Use inverted only with verified light/dark themes: the floor Teams high contrast theme resolves both inverted tokens to black. disabled is styling for an actually disabled element, not a native disabled attribute.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface needs coordinated Fluent foreground and background roles that can still be overridden property by property.
 * @example Basic usage
 * ```ts
 * import { presets } from '@apvee/spfx-react-toolkit/styles';
 *
 * const descriptor = presets.canvas;
 * ```
 *
 * @scenario Select paired color preset descriptors — foreground/background override composition
 * ```tsx
 * import * as React from 'react';
 * import { Checkbox, Field, Select } from '@fluentui/react-components';
 * import { useSx, width, minWidth, overflow, padding, presets, foreground, background, type SxBaseDescriptor } from '@apvee/spfx-react-toolkit';
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
 * const choices: Readonly<Record<string, SxBaseDescriptor>> = presets;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const [member, setMember] = React.useState('canvas');
 *   const [foregroundOverride, setForegroundOverride] = React.useState(false);
 *   const [backgroundOverride, setBackgroundOverride] = React.useState(false);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setMember('canvas'); setForegroundOverride(false); setBackgroundOverride(false);
 *   }, [resetRevision]);
 *   const descriptor = choices[member];
 *   return render({
 *     controls: <><Field label="Descriptor"><Select value={member} onChange={(_, data) => {
 *       if (Object.prototype.hasOwnProperty.call(choices, data.value)) setMember(data.value);
 *     }}>{Object.keys(choices).map(key => <option key={key} value={key}>{key}</option>)}</Select></Field>
 *       <Checkbox label="Override foreground with brand" checked={foregroundOverride} onChange={(_, data) => setForegroundOverride(data.checked === true)} />
 *       <Checkbox label="Override background with subtle" checked={backgroundOverride} onChange={(_, data) => setBackgroundOverride(data.checked === true)} />
 *     </>,
 *     status: <p aria-live="polite">Applied: presets.{member}. Foreground override: {foregroundOverride ? 'brand' : 'none'}. Background override: {backgroundOverride ? 'subtle' : 'none'}.</p>,
 *     children: <section role="region" aria-label="presets descriptor preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.auto, padding.medium)}>
 *       <div data-style-helper="presets" aria-disabled={member === 'disabled'} className={sx(padding.medium, descriptor,
 *         foregroundOverride && foreground.brand, backgroundOverride && background.subtle)}>Actual presets subject</div>
 *     </section>,
 *     diagnostics: <p>Later foreground.brand replaces only the preset foreground; later background.subtle replaces only its background. Toggle them independently, together, then off to restore each preset property. Inverted is intended for the verified light/dark palettes; it is unsuitable for active content in the installed Fluent high contrast palette. Token descriptors inherit the host theme; Reset restores canvas with both overrides off without remounting the Demo.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Select a paired preset, then toggle later foreground.brand and background.subtle overrides independently. Each overrides only its own property; disabling it restores that preset property. Reset restores canvas with both overrides off while retaining this Demo.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Applies colorNeutralBackground1 background and colorNeutralForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.canvas, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const canvas: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.canvas', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackground1, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground1, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorNeutralBackground2 background and colorNeutralForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.alternative, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const alternative: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.alternative', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackground2, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground1, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorSubtleBackground background and colorNeutralForeground2 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.subtle, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const subtle: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.subtle', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorSubtleBackground, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForeground2, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorTransparentBackground background; leaves color inherited.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.transparent, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const transparent: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.transparent', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorTransparentBackground, { property: 'backgroundColor', fallback: 'transparent' }),
]);

/** Applies colorBrandBackground background and colorNeutralForegroundOnBrand foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.brand, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const brand: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.brand', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorBrandBackground, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundOnBrand, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorBrandBackground2 background and colorBrandForeground2 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.brandTint, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const brandTint: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.brandTint', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorBrandBackground2, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorBrandForeground2, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorNeutralBackgroundInverted background and colorNeutralForegroundInverted foreground.
 * Supported for verified light/dark surfaces; the floor high-contrast theme has 1:1 contrast.
 * Use canvas/alternative for content in that high-contrast theme.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.inverted, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const inverted: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.inverted', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackgroundInverted, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundInverted, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorStatusSuccessBackground1 background and colorStatusSuccessForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.success, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const success: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.success', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorStatusSuccessBackground1, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorStatusSuccessForeground1, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorStatusWarningBackground1 background and colorStatusWarningForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.warning, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const warning: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.warning', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorStatusWarningBackground1, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorStatusWarningForeground1, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorStatusDangerBackground1 background and colorStatusDangerForeground1 foreground.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.danger, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const danger: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.danger', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorStatusDangerBackground1, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorStatusDangerForeground1, { property: 'color', fallback: 'inherit' }),
]);

/** Applies colorNeutralBackgroundDisabled background and colorNeutralForegroundDisabled foreground.
 * For actual disabled presentation only; light/dark floor pairs have low text contrast.
 * Does not disable the element or assign aria-disabled.
 * Requires scoped Fluent theme variables; adds no geometry or interaction states.
 * Verify contrast on the actual surface, especially for transparent/subtle/disabled.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, presets, typography } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(presets.disabled, typography.body1)}>Content</div>;
 * }
 * ```
 */
export const disabled: SxBaseDescriptor = /*#__PURE__*/ createRecipe('presets.disabled', [
  /*#__PURE__*/ createDeclaration('backgroundColor', tokens.colorNeutralBackgroundDisabled, { property: 'backgroundColor', fallback: 'transparent' }),
  /*#__PURE__*/ createDeclaration('color', tokens.colorNeutralForegroundDisabled, { property: 'color', fallback: 'inherit' }),
]);
