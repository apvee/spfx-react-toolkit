/**
 * Source-owned Declaration header recipe.
 * @packageDocumentation
 * @docPage helpers/styles
 * @example Declaration header recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
/**
 * Source-owned Declared namespace recipe.
 * @example Declared namespace recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export namespace Declared { export const value = 1; }
