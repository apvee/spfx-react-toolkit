// resize-observer.ts
// Utility hook to observe element size changes

import { useEffect } from 'react';
import type { ContainerSize } from '../core/types';

/**
 * Hook to observe element size changes using ResizeObserver
 * 
 * @param element - DOM element to observe
 * @param onResize - Callback when size changes
 */
export function useResizeObserver(
  element: HTMLElement | undefined,
  onResize: (size: ContainerSize | undefined) => void
): void {
  useEffect(() => {
    // If no element or ResizeObserver not supported, clear size
    if (!element || typeof ResizeObserver === 'undefined') {
      onResize(undefined);
      return;
    }
    
    // Create ResizeObserver
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        const { width, height } = entry.contentRect;
        onResize({ width, height });
      }
    });
    
    // Start observing
    observer.observe(element);
    
    // Cleanup on unmount
    return () => {
      observer.disconnect();
    };
  }, [element, onResize]);
}
