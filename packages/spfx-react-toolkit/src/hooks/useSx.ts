import { useMemo } from 'react';
import { useSxStyleEnvironment } from '../helpers/styles/renderer.internal';
import { resolveSxInputs } from '../helpers/styles/resolve.internal';
import type { SxFunction, SxOptions } from '../helpers/styles/types';

/**
 * Compose immutable style descriptors and existing class names in argument order.
 * Uses the existing Griffel renderer and Fluent direction; no SPFx provider is required.
 * Token descriptors need Fluent theme variables in the element's CSS scope.
 * @param options Optional effective direction override; keep it consistent with host classes.
 * @returns Ordinary class composer, stable while the renderer and direction stay unchanged.
 * No arguments or only falsy inputs return an empty string. Last declaration wins
 * per property and scope; external Griffel class strings preserve their input position.
 * Factories are cached per renderer; cached CSS rules persist for that renderer's lifetime.
 * @example
 * const sx = useSx();
 * return <div className={sx(width.px(240), condition && foreground.subtle)} />;
 */
export function useSx(options?: SxOptions): SxFunction {
  const { renderer, dir } = useSxStyleEnvironment(options);
  return useMemo(() => (...inputs) => resolveSxInputs({ renderer, dir }, inputs), [renderer, dir]);
}
