export { width, width as renamedWidth } from './runtime-barrel';
export * from './runtime-barrel';
/**
 * Source-owned Runtime binding recipe.
 * @example Runtime binding recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export * as aliasWidth from './width';
/**
 * Source-owned Erased binding recipe.
 * @example Erased binding recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export type * as TypeWidth from './width';
export * from './star-intermediate';
import type * as ImportedWidth from './width';
export { ImportedWidth };
import type { width as NamedImportWidth } from './runtime-barrel';
export { NamedImportWidth };
export { type width as NamedTypeWidth } from './runtime-barrel';
export type { width as OtherTypeWidth } from './runtime-barrel';
export * as starErased from './type-star';
export * as importErased from './import-erased';
export * as shadowed from './shadowed';
export * as duplicateReal from './duplicate-real';
export * as nested from './nested';
export * as empty from './empty';
export * as typeOnly from './types';
export * as ambient from './ambient';
export * as cycle from './cycle-a';
export * as cycleReal from './cycle-real-a';
export { Declared } from './declaration';
export { Control } from './control';
import * as privateNamespace from './private';
export const supportReference = () => privateNamespace;
export * as overloads from './overloads';
export * as imports from './imports';
export * as ambientFile from './ambient-file-user';
export type * as erased from './erased';
export * as classOnly from './class-only';
export * as enumOnly from './enum-only';
export * as uninitialized from './uninitialized';
export * as constEnumOnly from './const-enum-only';
export type * as TypeNested from './nested';
export * from './named-erased-barrel';
/**
 * Source-owned Nonnamespace binding recipe.
 * @example Nonnamespace binding recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export type { Control as TypeControl } from './control';
export * as mergedFunction from './merged-function';
export * as destructured from './destructured';
/**
 * A runtime namespace binding with a default expression.
 * @example Default binding recipe
 * ```tsx
 * import * as React from 'react';
 * import { settings } from '@fixture/namespace';
 * export function Demo(): React.ReactElement { return <div>{settings.default.actual}</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'The default settings value is visible.' };
 * ```
 */
export * as settings from './settings';
export * as erasedDefault from './erased-default';
