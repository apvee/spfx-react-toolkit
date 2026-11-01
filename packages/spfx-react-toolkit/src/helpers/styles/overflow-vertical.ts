/**
 * Horizontal/vertical axis overflow declarations without changing the other axis. CSS may compute visible to auto when the other axis is hidden or auto.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/**
 * Sets vertical overflow to visible; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 * export function ScrollArea(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.vertical.visible)}>Content</div>;
 * }
 * ```
 */
export const visible: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowY', 'visible', { property: 'overflowY', fallback: 'visible' });

/**
 * Sets vertical overflow to hidden; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 * export function ScrollArea(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.vertical.hidden)}>Content</div>;
 * }
 * ```
 */
export const hidden: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowY', 'hidden', { property: 'overflowY', fallback: 'visible' });

/**
 * Sets vertical overflow to auto; the other axis retains native CSS behavior.
 * CSS can compute visible to auto when the other axis is hidden or auto.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, overflow } from '@apvee/spfx-react-toolkit/styles';
 * export function ScrollArea(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(overflow.vertical.auto)}>Content</div>;
 * }
 * ```
 */
export const auto: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('overflowY', 'auto', { property: 'overflowY', fallback: 'visible' });
