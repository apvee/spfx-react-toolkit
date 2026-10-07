import { makeStyles } from '@griffel/core';
import type { GriffelRenderer } from '@griffel/core';

export type SxStyleFactory = ReturnType<typeof makeStyles<'root'>>;
export interface SxRendererCache {
  readonly bindings: Map<string, SxStyleFactory>;
  readonly assignments: Map<string, SxStyleFactory>;
}
// A factory captures its first renderer's salt. Never share initialized factories
// across renderers, even when values and logical declarations are identical.
const rendererCaches = new WeakMap<GriffelRenderer, SxRendererCache>();
export function getSxRendererCache(renderer: GriffelRenderer): SxRendererCache {
  let cache = rendererCaches.get(renderer);
  if (!cache) {
    cache = { bindings: new Map(), assignments: new Map() };
    rendererCaches.set(renderer, cache);
  }
  return cache;
}
