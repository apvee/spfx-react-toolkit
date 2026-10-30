import { useMemo } from 'react';
import { useSxStyleEnvironment } from '../helpers/styles/renderer.internal';
import { resolveSxInputs } from '../helpers/styles/resolve.internal';
import type { SxFunction, SxOptions } from '../helpers/styles/types';

/**
 * Composes immutable style descriptors and existing classes in argument order using the current Griffel renderer and Fluent direction. No SPFx provider is required. Token descriptors need Fluent theme variables in the element's CSS scope. Empty/falsy composition returns an empty string. Last declaration wins per property/scope; external Griffel classes retain their position. Factories are cached per renderer and CSS rules persist for its lifetime.
 * @param options - Optional effective direction override; otherwise uses Fluent context (LTR without context). Keep it consistent with host classes.
 * @returns Class composer, stable while renderer/direction stay unchanged.
 *
 * @example
 * ```tsx
 * import { useSx, width, foreground } from '@apvee/spfx-react-toolkit';
 *
 * function Surface() {
 *   const sx = useSx();
 *   return <div className={sx(width.px(240), foreground.subtle)}>Content</div>;
 * }
 * ```
 */
export function useSx(options?: SxOptions): SxFunction {
  const { renderer, dir } = useSxStyleEnvironment(options);
  return useMemo(() => (...inputs) => resolveSxInputs({ renderer, dir }, inputs), [renderer, dir]);
}
