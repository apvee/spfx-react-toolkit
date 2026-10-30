// useSPFxContainerSize.ts
// Hook for container size detection with SharePoint-aligned breakpoints

import { useMemo } from 'react';
import { getSPFxContainerSize } from '../helpers/spfx-container.helpers';
import { useSPFxContainerInfo } from './useSPFxContainerInfo';

/**
 * Toolkit width categories using exclusive upper boundaries: small <480, medium <640, large <1024, xLarge <1366, xxLarge <1920, otherwise xxxLarge. This is a container measurement; before observation zero width is small.
 *
 * @example
 * ```ts
 * const category: SPFxContainerSize = 'small';
 * ```
 */
export type SPFxContainerSize = 
  | 'small'     // <480px
  | 'medium'    // 480-639px
  | 'large'     // 640-1023px
  | 'xLarge'    // 1024-1365px
  | 'xxLarge'   // 1366-1919px
  | 'xxxLarge'; // >= 1920px

/**
 * Return type for useSPFxContainerSize hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const container: SPFxContainerSizeInfo = useSPFxContainerSize();
 * console.log(container.isSmall, container.width);
 * ```
 */
export interface SPFxContainerSizeInfo {
  /** Container size category (Fluent UI 9 aligned) */
  readonly size: SPFxContainerSize;
  
  /**
   * True for the small category (width <480px), including the zero-width unmeasured fallback.
   */
  readonly isSmall: boolean;
  
  /** Is medium container (480-639px) - mobile landscape, small tablets */
  readonly isMedium: boolean;
  
  /** Is large container (640-1023px) - tablets, single column */
  readonly isLarge: boolean;
  
  /** Is extra large container (1024-1365px) - laptop, desktop standard */
  readonly isXLarge: boolean;
  
  /** Is extra extra large container (1366-1919px) - large desktop, wide screen */
  readonly isXXLarge: boolean;
  
  /** Is extra extra extra large container (>= 1920px) - 4K, ultra-wide, multi-column */
  readonly isXXXLarge: boolean;
  
  /**
   * Observed container width in CSS pixels, or zero before/without a supported measurement.
   */
  readonly width: number;
  
  /**
   * Observed container height in CSS pixels, or zero before/without a supported measurement.
   */
  readonly height: number;
}

/**
 * Classifies the observed container width using toolkit thresholds: small <480, medium <640, large <1024, xLarge <1366, xxLarge <1920, otherwise xxxLarge. These measure container width, not viewport width. Before a supported measurement, width and height default to zero and the category is small. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Category, exact-category boolean flags and dimensions in CSS pixels.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function ContainerLabel() {
 *   const { size, width } = useSPFxContainerSize();
 *   return <span>{size}: {width}px</span>;
 * }
 * ```
 */
export function useSPFxContainerSize(): SPFxContainerSizeInfo {
  const { size: containerSize } = useSPFxContainerInfo();
  
  return useMemo(() => {
    const width = containerSize?.width ?? 0;
    const height = containerSize?.height ?? 0;
    const size = getSPFxContainerSize(width);
    
    return {
      size,
      isSmall: size === 'small',
      isMedium: size === 'medium',
      isLarge: size === 'large',
      isXLarge: size === 'xLarge',
      isXXLarge: size === 'xxLarge',
      isXXXLarge: size === 'xxxLarge',
      width,
      height,
    };
  }, [containerSize]);
}
