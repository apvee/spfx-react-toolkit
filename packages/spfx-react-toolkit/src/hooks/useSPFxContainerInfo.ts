// useSPFxContainerInfo.ts
// Hook to access container element and size

import { useMemo } from 'react';
import { useSPFxRuntimeActions, useSPFxRuntimeSelector } from '../core/state.internal';
import { useResizeObserver } from '../utils/resize-observer.internal';
import type { ContainerSize } from '../core/types';

/**
 * Return type for useSPFxContainerInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const container: SPFxContainerInfo = useSPFxContainerInfo();
 * console.log(container.size?.width);
 * ```
 */
export interface SPFxContainerInfo {
  /**
   * Provider container DOM element, or undefined when the host supplies none.
   */
  readonly element: HTMLElement | undefined;
  
  /**
   * Observed content-box dimensions in CSS pixels; undefined before observation or without an element/ResizeObserver.
   */
  readonly size: ContainerSize | undefined;
}

/**
 * Reads the provider container element and observes its content-box size using the element document's ResizeObserver. Size arrives asynchronously and is undefined before observation, when the element is absent, or when its document lacks ResizeObserver. Observation disconnects on element replacement/unmount; no polyfill is installed. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Container element and optional content-box width/height in CSS pixels.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function ContainerLabel() {
 *   const { size } = useSPFxContainerInfo();
 *   return <span>{size ? `${size.width} × ${size.height}` : 'Size unavailable'}</span>;
 * }
 * ```
 */
export function useSPFxContainerInfo(): SPFxContainerInfo {
  const element = useSPFxRuntimeSelector(state => state.containerEl);
  const size = useSPFxRuntimeSelector(state => state.containerSize);
  
  const { setContainerSize } = useSPFxRuntimeActions();
  useResizeObserver(element, setContainerSize);
  
  return useMemo(() => ({
    element,
    size,
  }), [element, size]);
}
