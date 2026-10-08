import * as React from 'react';
import { useIsomorphicLayoutEffect } from '@fluentui/react-utilities';
import { useStableCallback } from '../../hooks/useStableCallback';
import { useResizeObserver } from '../../utils/resize-observer.internal';
import { scanAutoSkeleton, validateAutoSkeletonOptions } from './auto-skeleton-geometry.internal';
import type { AutoSkeletonBone } from './auto-skeleton-geometry.internal';

type MeasurementOptions = {
  root: HTMLElement | undefined;
  content: HTMLElement | undefined;
  loading: boolean;
  children: React.ReactNode;
  minSize: number;
  ignoreSelectors: readonly string[];
};
const sameGeometry = (previous: readonly AutoSkeletonBone[], next: readonly AutoSkeletonBone[]): boolean =>
  previous.length === next.length && previous.every((bone, index) => {
    const other = next[index];
    return bone.x === other.x && bone.y === other.y && bone.width === other.width
      && bone.height === other.height && bone.radius === other.radius && bone.shape === other.shape;
  });

/** Per-instance DOM measurement and subscriptions; no geometry/storage registry. */
export function useAutoSkeleton({ root, content, loading, children, minSize, ignoreSelectors }: MeasurementOptions): readonly AutoSkeletonBone[] {
  const [bones, setBones] = React.useState<readonly AutoSkeletonBone[]>([]);
  const frame = React.useRef<number>();
  const measure = useStableCallback(() => {
    const next = root && content ? scanAutoSkeleton(content, root, { minSize, ignoreSelectors }) : [];
    setBones(previous => sameGeometry(previous, next) ? previous : next);
  });
  const schedule = useStableCallback(() => {
    const ownerWindow = root?.ownerDocument.defaultView;
    if (!root || !content || !ownerWindow || frame.current !== undefined) return;
    if (typeof ownerWindow.requestAnimationFrame !== 'function') { measure(); return; }
    frame.current = ownerWindow.requestAnimationFrame(() => { frame.current = undefined; measure(); });
  });
  useIsomorphicLayoutEffect(() => {
    if (content) validateAutoSkeletonOptions(content, minSize, ignoreSelectors);
  }, [content, minSize, ignoreSelectors]);
  useIsomorphicLayoutEffect(() => { measure(); }, [root, content, loading, children, minSize, ignoreSelectors, measure]);
  useResizeObserver(root, schedule);
  React.useEffect(() => {
    const ownerWindow = root?.ownerDocument.defaultView;
    if (!root || !content || !ownerWindow) return;
    const observer = typeof ownerWindow.MutationObserver === 'function' ? new ownerWindow.MutationObserver(records => {
      if (records.some(record => record.target !== content || record.type !== 'attributes'
        || !['class', 'inert', 'aria-hidden'].includes(record.attributeName ?? ''))) schedule();
    }) : undefined;
    observer?.observe(content, { attributes: true, characterData: true, childList: true, subtree: true });
    ownerWindow.addEventListener('resize', schedule);
    content.addEventListener('load', schedule, true);
    const fonts = content.ownerDocument.fonts;
    fonts?.addEventListener('loadingdone', schedule);
    return () => {
      if (frame.current !== undefined) { ownerWindow.cancelAnimationFrame(frame.current); frame.current = undefined; }
      observer?.disconnect();
      ownerWindow.removeEventListener('resize', schedule);
      content.removeEventListener('load', schedule, true);
      fonts?.removeEventListener('loadingdone', schedule);
    };
  }, [root, content, schedule]);
  return bones;
}
