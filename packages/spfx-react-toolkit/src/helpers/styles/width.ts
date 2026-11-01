/**
 * Width descriptors for automatic, containing-block or explicit pixel sizing.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface must fill its containing block, follow automatic sizing, or keep a deliberate fixed pixel width.
 * @example
 * ```ts
 * import { width } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = width.auto;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, toPixels } from './descriptor.internal';

/** Automatic width.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(width.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('width', 'auto', { property: 'width', fallback: 'auto' });

/** Full width of the containing block.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(width.full)}>Content</div>;
 * }
 * ```
 */
export const full: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('width', '100%', { property: 'width', fallback: 'auto' });

/** Creates an explicit pixel width.
 * @param value - Finite nonnegative pixel length, including zero.
 * @returns Immutable width descriptor.
 * @throws RangeError - When value is negative, NaN, or infinite.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, width } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(width.px(240))}>Content</div>;
 * }
 * ```
 */
export function px(value: number): SxBaseDescriptor {
  return createDeclaration('width', toPixels(value), { property: 'width', fallback: 'auto' });
}
