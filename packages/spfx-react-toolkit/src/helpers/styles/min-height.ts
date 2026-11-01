/**
 * Minimum height descriptors for shrinkable flex content or an explicit pixel minimum.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A flex child must shrink so a bounded descendant can scroll, or retain a deliberate minimum height.
 * @example
 * ```ts
 * import { minHeight } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = minHeight.zero;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets minHeight to 0px.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, minHeight } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(minHeight.zero)}>Content</div>;
 * }
 * ```
 */
export const zero: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('minHeight', '0px', { property: 'minHeight', fallback: 'auto' });

/** Creates an explicit pixel minHeight.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable minHeight descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, minHeight } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(minHeight.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('minHeight', toPixels(value), { property: 'minHeight', fallback: 'auto' });
}
