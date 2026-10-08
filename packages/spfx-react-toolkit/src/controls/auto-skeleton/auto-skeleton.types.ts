import type * as React from 'react';

/** Props for skeleton geometry derived from mounted, bounded HTML content.
 * Native div attributes and ref are forwarded to the root. Apply the content's
 * flex/grid layout inside children: an owned block wrapper is always present.
 * @example
 * <AutoSkeleton loading={isLoading}><ProfileCard /></AutoSkeleton>
 */
export type AutoSkeletonProps = Omit<React.ComponentPropsWithoutRef<'div'>,
  'children' | 'dangerouslySetInnerHTML'> & {
  /** Mounted content to measure. Absent/unmeasurable content has no inferred placeholder. */
  children: React.ReactNode;
  /** Marks the root busy; measurable content becomes inert beneath a skeleton overlay. */
  loading: boolean;
  /** Fluent animation; defaults to wave and respects the host's reduced-motion preference. */
  animation?: 'wave' | 'pulse';
  /** Minimum normalized width and height in pixels. Defaults to 8; must be finite and positive. */
  minSize?: number;
  /** CSS selectors excluding matching descendants and their entire subtrees.
   * Invalid selectors throw. data-skeleton-ignore also excludes a subtree.
   */
  ignoreSelectors?: readonly string[];
  /** Direct HTML injection is unsupported; provide rendered children instead. */
  dangerouslySetInnerHTML?: never;
};
