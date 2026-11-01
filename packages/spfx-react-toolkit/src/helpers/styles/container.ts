/**
 * Named inline-size containment for ancestor queries. Adds no dimensions or display; provide host geometry because containment affects intrinsic sizing.
 *
 * @packageDocumentation
 * @docPage helpers/styles
 * @useWhen An ancestor must supply the named size reference for responsive descendants; provide its layout geometry separately.
 * @example
 * ```ts
 * import { container } from '@apvee/spfx-react-toolkit/styles';
 * const descriptor = container.inlineSize;
 * ```
 */
import type { SxBaseDescriptor } from './types';
import { createDeclaration, createRecipe } from './descriptor.internal';

/**
 * Establishes the named `apvee-sx` inline-size query container for descendants.
 * Adds no width, height, display or spacing. Responsive descriptors measure the
 * nearest named ancestor's inline size, with inclusive 480/640/1024px thresholds.
 * An element never queries its own size; vertical writing modes measure height.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSx, container, responsive, flex } from '@apvee/spfx-react-toolkit/styles';
 * export function ResponsiveRows(): React.ReactElement {
 *   const sx = useSx();
 *   return <section className={sx(container.inlineSize)}>
 *     <div className={sx(flex.column, responsive.medium(flex.row))}>Content</div>
 *   </section>;
 * }
 * ```
 */
export const inlineSize: SxBaseDescriptor = /*#__PURE__*/ createRecipe('container.inlineSize', [
  /*#__PURE__*/ createDeclaration('containerType', 'inline-size', { property: 'containerType', fallback: 'normal' }),
  /*#__PURE__*/ createDeclaration('containerName', 'apvee-sx', { property: 'containerName', fallback: 'none' })
]);
