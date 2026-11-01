/**
 * Fluent shadow recipes; these do not add surface geometry or colors.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A surface needs Fluent elevation styling while its background, dimensions and padding remain separately controlled.
 * @example
 * ```ts
 * import { boxShadow } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = boxShadow.none;
 * ```
 */
import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration } from './descriptor.internal';

/** Applies box-shadow none without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, boxShadow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(boxShadow.none)}>Content</div>;
 * }
 * ```
 */
export const none: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', 'none', { property: 'boxShadow', fallback: 'none' });

/** Applies box-shadow var(--shadow4) without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, boxShadow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(boxShadow.small)}>Content</div>;
 * }
 * ```
 */
export const small: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', tokens.shadow4, { property: 'boxShadow', fallback: 'none' });

/** Applies box-shadow var(--shadow8) without changing geometry.
 * Token shadows require scoped Fluent theme variables.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, boxShadow } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(boxShadow.medium)}>Content</div>;
 * }
 * ```
 */
export const medium: SxBaseDescriptor = /*#__PURE__*/ createDeclaration('boxShadow', tokens.shadow8, { property: 'boxShadow', fallback: 'none' });
