import { tokens } from '@fluentui/react-theme';
import { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe, SxProperty, toPixels } from './descriptor.internal';

/** Exact Fluent spacing suffixes supported by the public catalog. */
export type SpacingSuffix = 'None' | 'XS' | 'S' | 'M' | 'L' | 'XL';

/** Creates canonical longhands using the separate Fluent block and inline token scales. */
export function createSpacing(
  name: string,
  blockProperties: readonly SxProperty[],
  inlineProperties: readonly SxProperty[],
  suffix: SpacingSuffix
): SxBaseDescriptor {
  return createRecipe(name, [
    ...blockProperties.map(property => createDeclaration(property, tokens[`spacingVertical${suffix}`], {
      property, fallback: '0px'
    })),
    ...inlineProperties.map(property => createDeclaration(property, tokens[`spacingHorizontal${suffix}`], {
      property, fallback: '0px'
    }))
  ]);
}

/** Creates explicit canonical pixel longhands with the shared finite nonnegative validation. */
export function createSpacingPixels(name: string, properties: readonly SxProperty[], value: number): SxBaseDescriptor {
  const pixels = toPixels(value);
  return createRecipe(name, properties.map(property => createDeclaration(property, pixels, {
    property, fallback: '0px'
  })));
}
