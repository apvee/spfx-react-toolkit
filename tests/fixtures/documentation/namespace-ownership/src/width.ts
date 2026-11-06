/**
 * Source-owned Width recipe.
 * @packageDocumentation
 * @docPage helpers/styles
 * @scenario Width recipe
 * ```tsx
 * import * as React from 'react';
 * export function Demo(): React.ReactElement { return <div>Visible namespace</div>; }
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject', observe: 'A namespace recipe is visible.' };
 * ```
 */
/**
 * Source-owned Ordinary recipe.
 * @example Ordinary recipe
 * ```ts
 * const ordinary = 42;
 * ```
 */
export const auto = 'auto';
/** Pixels.
 * @example Member recipe
 * ```ts
 * const pixels = 12;
 * ```
 */
export function px(value: number): string { return `${value}px`; }
