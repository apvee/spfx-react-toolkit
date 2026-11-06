/**
 * Container observation requires a WebPart host element and document ResizeObserver.
 *
 * @packageDocumentation
 * @docPage hooks/theming
 */
// useSPFxContainerInfo.ts
// Hook to access container element and size

import { useMemo } from 'react';
import { useSPFxRuntimeActions, useSPFxRuntimeSelector } from '../core/state.internal';
import { useResizeObserver } from '../utils/resize-observer.internal';
import type { ContainerSize } from '../core/types';

/**
 * Nearest provider container element and asynchronously observed content-box dimensions.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxContainerInfo, useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const container: SPFxContainerInfo = useSPFxContainerInfo();
 *   console.log(container.size?.width);
 *   return <span>{container.size?.width ?? 'Unmeasured'}</span>;
 * }
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
 * @example Read ContainerInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';
 *
 * function ContainerLabel() {
 *   const { size } = useSPFxContainerInfo();
 *   return <span>{size ? `${size.width} × ${size.height}` : 'Size unavailable'}</span>;
 * }
 * ```
 *
 * @useWhen a WebPart needs its host element and observed CSS-pixel dimensions.
 *
 * @scenario Observe the native container and tracked size
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxContainerInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxContainerInfo result">
 *         <dt>Host element</dt><dd>{info.element?.tagName ?? 'Unavailable'}</dd>
 *         <dt>Width</dt><dd>{info.size ? String(info.size.width) : 'Unmeasured'}</dd>
 *         <dt>Height</dt><dd>{info.size ? String(info.size.height) : 'Unmeasured'}</dd>
 *       </dl>,
 *     diagnostics: <p>Resize the real host element. Its owner document supplies ResizeObserver; size is unavailable until observation and without native support. Replacement and unmount disconnect observation.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual ContainerInfo values and explicit unavailable fields as the native host changes.' };
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
