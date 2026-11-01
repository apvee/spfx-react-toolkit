/**
 * Minimum width descriptors for shrinkable content or an explicit nonnegative pixel minimum.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A flex/grid child must shrink below its intrinsic content width, or retain a deliberate minimum width.
 * @example
 * ```ts
 * import { minWidth } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = minWidth.zero;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Explicit zero minimum width.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, minWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(minWidth.zero)}>Content</div>;
 * }
 * ```
 */
export const zero: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('minWidth', '0px', { property: 'minWidth', fallback: 'auto' });

/** Creates an explicit pixel width.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable minWidth descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, minWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(minWidth.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('minWidth', toPixels(value), { property: 'minWidth', fallback: 'auto' });
}
