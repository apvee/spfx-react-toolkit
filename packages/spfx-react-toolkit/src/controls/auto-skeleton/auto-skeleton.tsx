/**
 * Geometry-derived loading placeholders over mounted, bounded HTML content.
 *
 * @packageDocumentation
 * @docPage controls/auto-skeleton
 */
import * as React from 'react';
import { mergeClasses, Skeleton, SkeletonItem } from '@fluentui/react-components';
import { useIsomorphicLayoutEffect, useMergedRefs } from '@fluentui/react-utilities';
import { useAutoSkeletonStyles, useAutoSkeletonGeometryClasses } from './auto-skeleton.styles.internal';
import { useAutoSkeleton } from './use-auto-skeleton.internal';
import type { AutoSkeletonProps } from './auto-skeleton.types';
import type { AutoSkeletonBone } from './auto-skeleton-geometry.internal';

const emptySelectors: readonly string[] = [];
const emptyGeometry: readonly AutoSkeletonBone[] = [];
/**
 * Derive Fluent skeleton items from the live geometry of mounted HTML content.
 * Children stay mounted in a block wrapper. Real text rectangles and visual leaves
 * produce only circle/rectangle items; text is inset by up to 2px per side for spacing.
 * Rectangles retain Fluent's rounded corners when content has no measured radius;
 * positive content radii are preserved and circles use a 50% radius.
 * Animation defaults to Fluent wave; `animation="none"` uses static themed Fluent items
 * without shimmer, while preserving loading visibility, focus and aria-busy behavior.
 * No geometry cache or generated inline styles.
 * Loading content is hidden/inert only when measurable and native inert is supported.
 * Otherwise content remains available with aria-busy. Native inert and DOM geometry
 * are browser requirements; server imports are safe but server rendering has no overlay.
 * Bounded HTML only: portals, shadow DOM, transforms and nested clipping are unsupported.
 * @param props - Loading state, scan configuration and native div attributes.
 * @returns A div root containing mounted content and an optional decorative skeleton.
 * @example Loading a mounted profile
 * ```tsx
 * import * as React from 'react';
 * import { AutoSkeleton } from '@apvee/spfx-react-toolkit';
 * const isLoading = true;
 * const photoUrl = '/images/profile.png';
 * const name = 'Ada';
 * <AutoSkeleton loading={isLoading} animation="pulse">
 *   <section><img data-skeleton-shape="circle" src={photoUrl} alt="" /><p>{name}</p></section>
 * </AutoSkeleton>
 * ```
 *
 * @example Static placeholders
 * ```tsx
 * import * as React from 'react';
 * import { AutoSkeleton } from '@apvee/spfx-react-toolkit';
 * const showPlaceholder = true;
 * function ProfileCard(): React.ReactElement { return <article>Ada — project owner</article>; }
 * <AutoSkeleton loading={showPlaceholder} animation="none"><ProfileCard /></AutoSkeleton>
 * ```
 *
 * @useWhen Representative mounted HTML exists and a loading overlay should follow its measured geometry.
 * @avoidWhen Use explicit skeleton placeholders for cold empty data, virtualized content or unsupported transformed geometry.
 *
 *
 * @example Compose shape container and conditional ignore markers
 * ```tsx
 * import * as React from 'react';
 * import { Card, CardHeader, Text } from '@fluentui/react-components';
 * import { AutoSkeleton, useSx, flex, gap } from '@apvee/spfx-react-toolkit';
 * function ProfileLoading({ loading, showSummary, ignoreNotes }: { loading: boolean; showSummary: boolean; ignoreNotes: boolean }) {
 *   const sx = useSx();
 *   return <div>
 *     <AutoSkeleton loading={loading}>
 *       <Card data-skeleton-shape="rectangle"><CardHeader header={<Text>Profile</Text>} /><Text>Profile details</Text></Card>
 *     </AutoSkeleton>
 *     <AutoSkeleton loading={loading}>
 *       <div data-skeleton-container="" className={sx(flex.column, gap.medium)}>
 *         {showSummary && <p>Profile summary</p>}
 *         <p data-skeleton-ignore={ignoreNotes ? '' : undefined}>Notes excluded when requested</p>
 *       </div>
 *     </AutoSkeleton>
 *   </div>;
 * }
 * ```
 */
export const AutoSkeleton = React.forwardRef<HTMLDivElement, AutoSkeletonProps>((props, forwardedRef) => {
  const { loading, children, animation = 'wave', minSize = 8, ignoreSelectors = emptySelectors, className, ...nativeProps } = props;
  delete nativeProps.dangerouslySetInnerHTML;
  const [root, setRoot] = React.useState<HTMLDivElement>();
  const [content, setContent] = React.useState<HTMLDivElement>();
  const rootCallback = React.useCallback((element: HTMLDivElement | null) => setRoot(element ?? undefined), []);
  const contentCallback = React.useCallback((element: HTMLDivElement | null) => setContent(element ?? undefined), []);
  const rootRef = useMergedRefs(rootCallback, forwardedRef);
  const styles = useAutoSkeletonStyles();
  const isStatic = animation === 'none';
  const bones = useAutoSkeleton({ root, content, loading, children, minSize, ignoreSelectors });
  const active = loading && bones.length > 0 && content !== undefined && 'inert' in content;
  const geometryClasses = useAutoSkeletonGeometryClasses(active ? bones : emptyGeometry);
  useIsomorphicLayoutEffect(() => {
    if (!active || !content) return;
    const focused = content.ownerDocument.activeElement;
    if (focused && content.contains(focused) && 'blur' in focused && typeof focused.blur === 'function') focused.blur();
  }, [active, content]);
  return <div {...nativeProps} ref={rootRef} aria-busy={loading} className={mergeClasses(styles.root, className)}>
    <div ref={contentCallback} data-auto-skeleton-content="" className={active ? styles.contentHidden : undefined}
      aria-hidden={active ? true : undefined} {...{ inert: active ? '' : undefined }}>{children}</div>
    {active && <Skeleton data-auto-skeleton-overlay="" animation={isStatic ? 'wave' : animation} aria-hidden="true" aria-busy={undefined}
      role="presentation" className={styles.overlay}>
      {bones.map((bone, index) => <SkeletonItem key={`${bone.x}:${bone.y}:${bone.width}:${bone.height}:${bone.radius}:${bone.shape}`}
        data-auto-skeleton-shape={bone.shape} shape={bone.shape} className={mergeClasses(styles.item, isStatic && styles.staticItem, geometryClasses[index])} />)}
    </Skeleton>}
  </div>;
});
AutoSkeleton.displayName = 'AutoSkeleton';
