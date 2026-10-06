// useSPFxContainerInfo.ts
// Hook to access container element and size

import { useMemo } from 'react';
import { useSPFxRuntimeActions, useSPFxRuntimeSelector } from '../core/state.internal';
import { useResizeObserver } from '../utils/resize-observer.internal';
import type { ContainerSize } from '../core/types';

/**
 * Return type for useSPFxContainerInfo hook
 */
export interface SPFxContainerInfo {
  /** Container DOM element */
  readonly element: HTMLElement | undefined;
  
  /** Container size (width/height in pixels) */
  readonly size: ContainerSize | undefined;
}

/**
 * Hook to access container element and its size
 * 
 * Automatically tracks container size changes using ResizeObserver
 * Updates in real-time when container is resized
 * 
 * Useful for:
 * - Responsive layouts
 * - Dynamic content sizing
 * - Breakpoint calculations
 * - Canvas/chart dimensions
 * 
 * @returns Container element and size
 * 
 * @see {@link useSPFxContainerSize} for responsive breakpoint categories (small/medium/large)
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { element, size } = useSPFxContainerInfo();
 *   
 *   return (
 *     <div>
 *       {size && (
 *         <p>Container: {size.width}px × {size.height}px</p>
 *       )}
 *     </div>
 *   );
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
