// useSPFxContainerSize.ts
// Hook for container size detection with SharePoint-aligned breakpoints

import { useMemo } from 'react';
import { getSPFxContainerSize } from '../helpers/spfx-container.helpers';
import { useSPFxContainerInfo } from './useSPFxContainerInfo';

/**
 * Container size categories (Fluent UI 9 aligned)
 * 
 * Based on official Fluent UI 9 responsive breakpoints:
 * @see https://developer.microsoft.com/en-us/fluentui#/styles/web/responsive
 * 
 * - small:    320-479px   (Mobile portrait)
 * - medium:   480-639px   (Mobile landscape, small tablets)
 * - large:    640-1023px  (Tablets, single column)
 * - xLarge:   1024-1365px (Laptop, desktop standard)
 * - xxLarge:  1366-1919px (Large desktop, wide screen)
 * - xxxLarge: >= 1920px   (4K, ultra-wide, multi-column)
 */
export type SPFxContainerSize = 
  | 'small'     // 320-479px
  | 'medium'    // 480-639px
  | 'large'     // 640-1023px
  | 'xLarge'    // 1024-1365px
  | 'xxLarge'   // 1366-1919px
  | 'xxxLarge'; // >= 1920px

/**
 * Return type for useSPFxContainerSize hook
 */
export interface SPFxContainerSizeInfo {
  /** Container size category (Fluent UI 9 aligned) */
  readonly size: SPFxContainerSize;
  
  /** Is small container (320-479px) - mobile portrait */
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
  
  /** Actual container width in pixels */
  readonly width: number;
  
  /** Actual container height in pixels */
  readonly height: number;
}

/**
 * Hook to get container size with Fluent UI 9 aligned breakpoints
 * 
 * Automatically tracks container size changes and categorizes into
 * Fluent UI 9 compatible size categories.
 * 
 * Breakpoints (aligned with Fluent UI 9):
 * - small:    320-479px   (mobile portrait)
 * - medium:   480-639px   (mobile landscape, small tablets)
 * - large:    640-1023px  (tablets, single column)
 * - xLarge:   1024-1365px (laptop, desktop standard)
 * - xxLarge:  1366-1919px (large desktop, wide screen)
 * - xxxLarge: >= 1920px   (4K, ultra-wide, multi-column)
 * 
 * Useful for:
 * - Responsive layouts based on actual container space
 * - Adaptive UI that works in sidebars, columns, or full width
 * - Fluent UI 9 compatible responsive design
 * - Making layout decisions based on available space
 * 
 * @returns Container size information and helpers
 * 
 * @example
 * ```tsx
 * function MyWebPart() {
 *   const { size, isSmall, isXXXLarge, width } = useSPFxContainerSize();
 *   
 *   // Decision based on size
 *   if (size === 'small') {
 *     return <CompactMobileView />;
 *   }
 *   
 *   if (size === 'medium' || size === 'large') {
 *     return <TabletView />;
 *   }
 *   
 *   if (size === 'xxxLarge') {
 *     return <UltraWideView columns={6} />;  // 4K/ultra-wide
 *   }
 *   
 *   return <DesktopView columns={size === 'xxLarge' ? 4 : 3} />;
 * }
 * ```
 * 
 * @example
 * ```tsx
 * function ResponsiveComponent() {
 *   const { isSmall, isMedium, isLarge, isXLarge, isXXLarge, isXXXLarge } = useSPFxContainerSize();
 *   
 *   return (
 *     <div>
 *       {isSmall && <MobileLayout />}
 *       {(isMedium || isLarge) && <TabletLayout />}
 *       {isXLarge && <DesktopLayout columns={3} />}
 *       {isXXLarge && <DesktopLayout columns={4} />}
 *       {isXXXLarge && <UltraWideLayout columns={6} />}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example
 * ```tsx
 * function AdaptiveCard() {
 *   const { size, width, height } = useSPFxContainerSize();
 *   
 *   return (
 *     <div>
 *       <p>Container: {width}px × {height}px</p>
 *       <p>Size category: {size}</p>
 *       {size === 'small' && <StackedLayout />}
 *       {size === 'medium' && <TwoColumnLayout />}
 *       {size === 'large' && <ThreeColumnLayout />}
 *       {(size === 'xLarge' || size === 'xxLarge') && <MultiColumnLayout />}
 *       {size === 'xxxLarge' && <UltraWideLayout columns={6} />}
 *     </div>
 *   );
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
