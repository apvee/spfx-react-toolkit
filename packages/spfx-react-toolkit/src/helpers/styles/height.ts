/**
 * Height descriptors; percentage height requires a definite containing-block height.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A region needs an explicit height; use full only when its containing block has a definite height.
 * @example
 * ```ts
 * import { height } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = height.auto;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets height to auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, height } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(height.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('height', 'auto', { property: 'height', fallback: 'auto' });

/** Full height of the containing block; percentage sizing requires a definite containing height.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, height } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(height.full)}>Content</div>;
 * }
 * ```
 */
export const full: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('height', '100%', { property: 'height', fallback: 'auto' });

/** Creates an explicit pixel height.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable height descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, height } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(height.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('height', toPixels(value), { property: 'height', fallback: 'auto' });
}
