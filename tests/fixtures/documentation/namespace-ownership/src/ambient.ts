/**
 * Source-owned Ambient recipe.
 * @packageDocumentation
 * @docPage helpers/styles
 * @scenario Ambient recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export declare function actual(): number;
export declare const other: number;
