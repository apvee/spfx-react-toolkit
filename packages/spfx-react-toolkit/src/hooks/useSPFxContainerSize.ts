/**
 * Width categories derive from the measured container, with zero as the unmeasured fallback.
 *
 * @packageDocumentation
 * @docPage hooks/theming
 */
// useSPFxContainerSize.ts
// Hook for container size detection with SharePoint-aligned breakpoints

import { useMemo } from 'react';
import { getSPFxContainerSize } from '../helpers/spfx-container.helpers';
import { useSPFxContainerInfo } from './useSPFxContainerInfo';

/**
 * Toolkit width categories using exclusive upper boundaries: small `less than 480`, medium `less than 640`, large `less than 1024`, xLarge `less than 1366`, xxLarge `less than 1920`, otherwise xxxLarge. This is a container measurement; before observation zero width is small.
 *
 * @example
 * ```ts
 * import { SPFxContainerSize } from '@apvee/spfx-react-toolkit';
 *
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
 * Exact toolkit container-width category, category flags and observed CSS-pixel dimensions.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxContainerSizeInfo, useSPFxContainerSize } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const container: SPFxContainerSizeInfo = useSPFxContainerSize();
 *   console.log(container.isSmall, container.width);
 *   return <span>{container.width}</span>;
 * }
 * ```
 */
export interface SPFxContainerSizeInfo {
  /** Container size category (Fluent UI 9 aligned) */
  readonly size: SPFxContainerSize;
  
  /**
   * True for the small category (width below 480px), including the zero-width unmeasured fallback.
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
  
  /** Is extra extra extra large container (1920px or wider) - 4K, ultra-wide, multi-column */
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
 * Classifies the observed container width using toolkit thresholds: small `less than 480`, medium `less than 640`, large `less than 1024`, xLarge `less than 1366`, xxLarge `less than 1920`, otherwise xxxLarge. These measure container width, not viewport width. Before a supported measurement, width and height default to zero and the category is small. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Category, exact-category boolean flags and dimensions in CSS pixels.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read ContainerSize metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContainerSize } from '@apvee/spfx-react-toolkit';
 *
 * function ContainerLabel() {
 *   const { size, width } = useSPFxContainerSize();
 *   return <span>{size}: {width}px</span>;
 * }
 * ```
 *
 * @useWhen a WebPart layout must switch on measured container width categories.
 *
 * @example Observe the measured container width category
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContainerSize } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxContainerSize();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxContainerSize result">
 *         <dt>Width</dt><dd>{String(info.width)}</dd>
 *         <dt>Height</dt><dd>{String(info.height)}</dd>
 *         <dt>Category</dt><dd>{info.size}</dd>
 *         <dt>Small</dt><dd>{String(info.isSmall)}</dd>
 *         <dt>Medium</dt><dd>{String(info.isMedium)}</dd>
 *         <dt>Large</dt><dd>{String(info.isLarge)}</dd>
 *         <dt>Extra large</dt><dd>{String(info.isXLarge)}</dd>
 *         <dt>Double extra large</dt><dd>{String(info.isXXLarge)}</dd>
 *         <dt>Triple extra large</dt><dd>{String(info.isXXXLarge)}</dd>
 *       </dl>,
 *     diagnostics: <p>Dimensions measure the host container in CSS pixels. Zero/small is the unmeasured fallback. Resize the actual host; no example control substitutes dimensions.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual ContainerSize values and explicit unavailable fields as the native host changes.' };
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
