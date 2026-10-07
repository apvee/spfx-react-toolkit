import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/** Creates a grid with equally sized columns that may shrink below content width.
 * @param count Positive integer column count.
 * @returns Immutable display and column-template recipe.
 * @throws {RangeError} When count is not a finite positive integer.
 * @example sx(grid.columns(3), gap.medium)
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
