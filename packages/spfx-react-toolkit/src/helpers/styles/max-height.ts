/**
 * Maximum height descriptors; none removes a maximum constraint.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A content region needs a height cap, optionally combined with overflow scrolling.
 * @example
 * ```ts
 * import { maxHeight } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = maxHeight.none;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Sets maxHeight to none.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, maxHeight } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(maxHeight.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('maxHeight', 'none', { property: 'maxHeight', fallback: 'none' });

/** Creates an explicit pixel maxHeight.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable maxHeight descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, maxHeight } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(maxHeight.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('maxHeight', toPixels(value), { property: 'maxHeight', fallback: 'none' });
}
