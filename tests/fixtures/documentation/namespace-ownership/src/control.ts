/**
 * Source-owned Control header recipe.
 * @packageDocumentation
 * @docPage helpers/styles
 * @example Control header recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
/**
 * Source-owned Control declaration recipe.
 * @example Control declaration recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export const Control = 1;
