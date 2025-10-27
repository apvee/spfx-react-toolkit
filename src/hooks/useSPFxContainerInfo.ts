// useSPFxContainerInfo.ts
// Hook to access container element and size

import { useAtomValue, useSetAtom } from 'jotai';
import { useSPFxContext } from './useSPFxContext';
import { spfxAtoms } from '../core/atoms';
import { useResizeObserver } from '../utils/resize-observer';
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
  const { instanceId } = useSPFxContext();
  
  // Get atoms for this instance
  const containerElAtom = spfxAtoms.containerEl(instanceId);
  const containerSizeAtom = spfxAtoms.containerSize(instanceId);
  
  // Read current values
  const element = useAtomValue(containerElAtom);
  const size = useAtomValue(containerSizeAtom);
  
  // Setup ResizeObserver with atom setter
  const setSize = useSetAtom(containerSizeAtom);
  useResizeObserver(element, setSize);
  
  return {
    element,
    size,
  };
}
