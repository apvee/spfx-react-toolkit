/**
 * Maximum width descriptors; none removes a maximum constraint.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen Long content needs a width cap while its preferred width continues to follow layout.
 * @example
 * ```ts
 * import { maxWidth } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = maxWidth.none;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets maxWidth to none.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, maxWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(maxWidth.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('maxWidth', 'none', { property: 'maxWidth', fallback: 'none' });

/** Creates an explicit pixel maxWidth.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable maxWidth descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, maxWidth } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(maxWidth.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('maxWidth', toPixels(value), { property: 'maxWidth', fallback: 'none' });
}
