/**
 * Flex grow/shrink declarations for items in a flex container.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An item in an existing flex container must share spare space or control shrinking under pressure.
 * @example
 * ```ts
 * import { flexItem } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = flexItem.grow;
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Sets flexGrow to 1 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.grow)}>Content</div>;
 * }
 * ```
 */
export const grow: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexGrow', '1', { property: 'flexGrow', fallback: '0' });

/** Sets flexGrow to 0 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.noGrow)}>Content</div>;
 * }
 * ```
 */
export const noGrow: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexGrow', '0', { property: 'flexGrow', fallback: '0' });

/** Sets flexShrink to 1 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.shrink)}>Content</div>;
 * }
 * ```
 */
export const shrink: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexShrink', '1', { property: 'flexShrink', fallback: '1' });

/** Sets flexShrink to 0 on a flex item.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, flexItem } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(flexItem.noShrink)}>Content</div>;
 * }
 * ```
 */
export const noShrink: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('flexShrink', '0', { property: 'flexShrink', fallback: '1' });
