/**
 * Source-owned Type-only recipe.
 * @packageDocumentation
 * @docPage helpers/styles
 * @scenario Type-only recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export interface Options { value: number; }
