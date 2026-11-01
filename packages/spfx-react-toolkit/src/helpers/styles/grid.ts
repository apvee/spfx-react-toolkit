/**
 * Equal-width grid columns with minmax(0, 1fr) tracks and explicit positive integer counts.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen A collection needs a known positive number of equally sized columns that may shrink below intrinsic content width.
 * @example
 * ```ts
 * import { grid } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = grid.columns(3);
 * ```
 */
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Creates a grid with equally sized columns that may shrink below content width.
 * @param count - Positive integer column count.
 * @returns Immutable display and column-template recipe.
 * @throws RangeError - When count is not a finite positive integer.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, grid, gap } from '@apvee/spfx-react-toolkit/styles';
 *
 * export function StyledSurface(): React.ReactElement {
 *   const sx = useSx();
 *   return <div className={sx(grid.columns(3), gap.medium)}>Content</div>;
 * }
 * ```
 */
export function columns(count: number): SxBaseDescriptor {
  if (!Number.isInteger(count) || count <= 0) {
    throw new RangeError('Grid column count must be a positive integer.');
  }
  // CSS repeat() requires decimal integer syntax, while Number stringification may use an exponent.
  const integer = count.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 0 });
  return createRecipe('grid.columns', [
    createDeclaration('display', 'grid', { property: 'display', fallback: 'inline' }),
    createDeclaration('gridTemplateColumns', `repeat(${integer}, minmax(0, 1fr))`, {
      property: 'gridTemplateColumns', fallback: 'none'
    })
  ]);
}
