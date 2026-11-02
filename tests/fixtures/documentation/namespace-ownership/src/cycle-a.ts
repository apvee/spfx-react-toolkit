/**
 * Source-owned Cycle recipe.
 * @packageDocumentation
 * @docPage helpers/styles
 * @example Cycle recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
export * as next from './cycle-b';
