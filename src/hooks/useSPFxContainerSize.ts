// useSPFxContainerSize.ts
// Hook for container size detection with SharePoint-aligned breakpoints

import { useMemo } from 'react';
import { useSPFxContainerInfo } from './useSPFxContainerInfo';

/**
 * Container size categories (SharePoint/Fluent UI aligned)
 * 
 * Based on official SharePoint Grid & Responsive Design breakpoints:
 * @see https://learn.microsoft.com/sharepoint/dev/design/grid-and-responsive-design
 * 
 * - small:   < 480px   (Mobile portrait, narrow sidebar)
 * - medium:  480-640px (Mobile landscape, small tablets)
 * - large:   640-1024px (Tablets, single column)
 * - xLarge:  1024-1366px (Laptop, desktop standard)
 * - xxLarge: > 1366px  (Large desktop, 4K, multi-column)
 */
export type SPFxContainerSize = 
  | 'small'    // < 480px
  | 'medium'   // 480-640px
  | 'large'    // 640-1024px
  | 'xLarge'   // 1024-1366px
  | 'xxLarge'; // > 1366px

/**
 * Return type for useSPFxContainerSize hook
 */
export interface SPFxContainerSizeInfo {
  /** Container size category (SharePoint aligned) */
  readonly size: SPFxContainerSize;
  
  /** Is small container (< 480px) - mobile portrait, narrow sidebar */
  readonly isSmall: boolean;
  
  /** Is medium container (480-640px) - mobile landscape, small tablets */
  readonly isMedium: boolean;
  
  /** Is large container (640-1024px) - tablets, single column */
  readonly isLarge: boolean;
  
  /** Is extra large container (1024-1366px) - laptop, desktop standard */
  readonly isXLarge: boolean;
  
  /** Is extra extra large container (> 1366px) - large desktop, multi-column */
  readonly isXXLarge: boolean;
  
  /** Actual container width in pixels */
  readonly width: number;
  
  /** Actual container height in pixels */
  readonly height: number;
}

/**
 * SharePoint container size breakpoints
 * 
 * Based on official SharePoint responsive grid system:
 * - Small:   320px  (1 column, no gutter)
 * - Medium:  480px  (12 columns, 16px gutter)
 * - Large:   640px  (12 columns, 24px gutter)
 * - XLarge:  1024px (12 columns, 24px gutter)
 * - XXLarge: 1366px (12 columns, 32px gutter)
 * 
 * @see https://learn.microsoft.com/sharepoint/dev/design/grid-and-responsive-design#breakpoints
 */
const CONTAINER_SIZE_BREAKPOINTS = {
  small: 480,    // SharePoint Small → Medium
  medium: 640,   // SharePoint Medium → Large
  large: 1024,   // SharePoint Large → XLarge
  xLarge: 1366,  // SharePoint XLarge → XXLarge
  // xxLarge: > 1366px (SharePoint XXLarge+)
} as const;

/**
 * Get container size category from width
 * 
 * @param width - Container width in pixels
 * @returns Container size category
 */
function getContainerSize(width: number): SPFxContainerSize {
  if (width < CONTAINER_SIZE_BREAKPOINTS.small) {
    return 'small';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.medium) {
    return 'medium';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.large) {
    return 'large';
  }
  if (width < CONTAINER_SIZE_BREAKPOINTS.xLarge) {
    return 'xLarge';
  }
  return 'xxLarge';
}

/**
 * Hook to get container size with SharePoint-aligned breakpoints
 * 
 * Automatically tracks container size changes and categorizes into
 * SharePoint/Fluent UI compatible size categories.
 * 
 * Breakpoints (aligned with SharePoint):
 * - small:   < 480px   (mobile portrait, narrow sidebar)
 * - medium:  480-640px (mobile landscape, small tablets)
 * - large:   640-1024px (tablets, single column)
 * - xLarge:  1024-1366px (laptop, desktop standard)
 * - xxLarge: > 1366px  (large desktop, multi-column)
 * 
 * Useful for:
 * - Responsive layouts based on actual container space
 * - Adaptive UI that works in sidebars, columns, or full width
 * - SharePoint-compatible responsive design
 * - Making layout decisions based on available space
 * 
 * @returns Container size information and helpers
 * 
 * @example
 * ```tsx
 * function MyWebPart() {
 *   const { size, isSmall, isXLarge, width } = useSPFxContainerSize();
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
 *   return <DesktopView columns={size === 'xxLarge' ? 4 : 3} />;
 * }
 * ```
 * 
 * @example
 * ```tsx
 * function ResponsiveComponent() {
 *   const { isSmall, isMedium, isLarge, isXLarge, isXXLarge } = useSPFxContainerSize();
 *   
 *   return (
 *     <div>
 *       {isSmall && <MobileLayout />}
 *       {(isMedium || isLarge) && <TabletLayout />}
 *       {isXLarge && <DesktopLayout columns={3} />}
 *       {isXXLarge && <DesktopLayout columns={4} />}
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
    const size = getContainerSize(width);
    
    return {
      size,
      isSmall: size === 'small',
      isMedium: size === 'medium',
      isLarge: size === 'large',
      isXLarge: size === 'xLarge',
      isXXLarge: size === 'xxLarge',
      width,
      height,
    };
  }, [containerSize]);
}
