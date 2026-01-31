// resize-observer.ts
// Utility hook to observe element size changes

import { useEffect } from 'react';
import type { ContainerSize } from '../core/types';

/**
 * Hook to observe element size changes using ResizeObserver
 *
 * Automatically sets up and cleans up ResizeObserver subscription.
 * Calls onResize with undefined if element is undefined or ResizeObserver is not supported.
 *
 * @param element - DOM element to observe, or undefined to clear observation
 * @param onResize - Callback invoked with {width, height} on size changes, or undefined when element is unavailable
 * @returns void - Hook manages subscription lifecycle internally
 *
 * @remarks
 * - Cleanup is handled automatically on unmount or when element changes
 * - ResizeObserver polyfill is NOT included - requires browser support
 * - Initial size is reported immediately after observation starts
 *
 * @internal
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
