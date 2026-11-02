/**
 * Fluent scrollbar treatment with native/forced-color fallbacks. Adds no overflow or dimensions; verify thumb contrast against the actual underlying surface.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An already-scrollable region needs compact theme-aware scrollbar styling with native forced-color fallback.
 * @example Basic usage
 * ```ts
 * import { scrollbar } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = scrollbar.fluent;
 * ```
 *
 * @example Scroll with the Fluent scrollbar descriptor — both axes and native keyboard/drag
 * ```tsx
 * import * as React from 'react';
 * import { Checkbox, makeStyles } from '@fluentui/react-components';
 * import { useSx, width, height, maxWidth, overflow, padding, presets, scrollbar } from '@apvee/spfx-react-toolkit';
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
 * const useProbeStyles = makeStyles({ surface: { boxSizing: 'border-box' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement;
 *   resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx();
 *   const classes = useProbeStyles();
 *   const [enabled, setEnabled] = React.useState(true);
 *   const revision = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (revision.current === resetRevision) return;
 *     revision.current = resetRevision;
 *     setEnabled(true);
 *   }, [resetRevision]);
 *   return render({
 *     controls: <Checkbox label="Fluent scrollbar" checked={enabled} onChange={(_, data) => setEnabled(data.checked === true)} />,
 *     status: <p>Scroll both axes with the keyboard or drag their native thumbs. Chromium has 6px rounded feedback; other engines retain native thin fallback.</p>,
 *     children: <div data-style-helper="scrollbar" role="region" aria-label="Native scrollbar preview" tabIndex={0}
 *       className={sx(classes.surface, height.px(140), width.full, maxWidth.px(320), overflow.auto, enabled && scrollbar.fluent, presets.alternative, padding.small)}>
 *       <div className={sx(width.px(600))}>{Array.from({ length: 14 }, (_, index) => <p key={index}>Scrollable row {index + 1}</p>)}</div>
 *     </div>,
 *     diagnostics: <p>Forced colors retains native/system colors and forcedColorAdjust:auto. Reset re-enables the descriptor without replacing the renderer or the scroll surface.</p>,
 *   });
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'configuration',
 *   observe: 'Toggle the real Fluent scrollbar descriptor on a bounded two-axis overflowing surface; use native arrows, thumb hover and dragging. Reset re-enables the descriptor. CSS rules alone do not prove native scrollbar paint.',
 * };
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

const thumbColor = tokens.colorNeutralStrokeAccessible;

/** Fluent scrollbar appearance: 6px on both axes in browsers supporting WebKit
 * scrollbar pseudo-elements (including Edge/Chrome), otherwise native thin width;
 * accessible neutral stroke thumb with 3px rounded corners and transparent track.
 * The thumb uses colorNeutralStrokeAccessibleHover on hover and
 * colorNeutralStrokeAccessiblePressed while pressed/dragged; pressed wins while
 * hovered. Forced colors use the browser's automatic colors and interactions.
 * Requires scoped theme variables. Vendor styling applies only outside forced
 * colors; standard scrollbar-width/color support supplies the native fallback.
 * Adds no overflow, container dimensions, gutter, overscroll, or smooth behavior.
 * Verify thumb contrast against the actual underlying surface.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, scrollbar, overflow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.auto, scrollbar.fluent)}>Content</div>;
 * }
 * ```
 */
export const fluent: SxBaseDescriptor = /*#__PURE__*/ createRecipe('scrollbar.fluent', [
  /*#__PURE__*/ createDeclaration('scrollbarWidth', 'thin', {
    property: 'scrollbarWidth', fallback: 'auto', scrollbarSize: '6px'
  }),
  /*#__PURE__*/ createDeclaration('scrollbarColor', `${thumbColor} transparent`, {
    property: 'scrollbarColor', fallback: 'auto', forcedColors: 'auto', scrollbarThumb: thumbColor,
    scrollbarThumbHover: tokens.colorNeutralStrokeAccessibleHover,
    scrollbarThumbPressed: tokens.colorNeutralStrokeAccessiblePressed
  })
]);
