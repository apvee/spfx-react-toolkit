import type { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/**
 * Establishes the named `apvee-sx` inline-size query container for descendants.
 * Adds no width, height, display or spacing. Responsive descriptors measure the
 * nearest named ancestor's inline size, with inclusive 480/640/1024px thresholds.
 * An element never queries its own size; vertical writing modes measure height.
 * @example <section className={sx(container.inlineSize)}><div className={sx(responsive.medium(width.full))} /></section>
 */
export const inlineSize: SxBaseDescriptor = /*#__PURE__*/ createRecipe('container.inlineSize', [
  /*#__PURE__*/ createDeclaration('containerType', 'inline-size', { property: 'containerType', fallback: 'normal' }),
  /*#__PURE__*/ createDeclaration('containerName', 'apvee-sx', { property: 'containerName', fallback: 'none' })
]);
