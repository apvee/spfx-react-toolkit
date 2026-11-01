/**
 * Horizontal/vertical axis overflow declarations without changing the other axis. CSS may compute visible to auto when the other axis is hidden or auto.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/**
 * Sets horizontal overflow to visible; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 * export function ScrollArea(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.horizontal.visible)}>Content</div>;
 * }
 * ```
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowX', 'visible', { property: 'overflowX', fallback: 'visible' });

/**
 * Sets horizontal overflow to hidden; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 * export function ScrollArea(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.horizontal.hidden)}>Content</div>;
 * }
 * ```
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowX', 'hidden', { property: 'overflowX', fallback: 'visible' });

/**
 * Sets horizontal overflow to auto; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 * export function ScrollArea(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.horizontal.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowX', 'auto', { property: 'overflowX', fallback: 'visible' });
