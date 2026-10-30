import type * as React from 'react';

/** Props for skeleton geometry derived from mounted, bounded HTML content.
 * Native div attributes and ref are forwarded to the root. Uses the host Fluent theme
 * and Griffel renderer; no SPFx provider is required. Native inert and bounded browser DOM
 * geometry are required for the overlay; server rendering keeps only the content. Apply the content's
 * flex/grid layout inside children: an owned block wrapper is always present.
 * @example
 * ```tsx
 * <AutoSkeleton loading={isLoading}><ProfileCard /></AutoSkeleton>
 * ```
 */
export type AutoSkeletonProps = Omit<React.ComponentPropsWithoutRef<'div'>,
  'children' | 'dangerouslySetInnerHTML'> & {
  /** Mounted content to measure. Absent/unmeasurable content has no inferred placeholder. */
  children: React.ReactNode;
  /** Sets root `aria-busy`. When geometry and native `inert` exist, hides measured content
   * beneath a decorative overlay and releases focus inside it. Otherwise content stays available. */
  loading: boolean;
  /** Visual treatment; defaults to `wave`. Fluent `wave` and `pulse` respect reduced motion.
   * `none` keeps Fluent items with a uniform themed fill and no animation or shimmer.
   * This affects presentation only: `loading` still controls visibility, inert content and `aria-busy`.
   */
  animation?: 'wave' | 'pulse' | 'none';
  /** Minimum normalized width and height in pixels. Defaults to `8`; nonfinite or nonpositive values throw `RangeError` after mount. */
  minSize?: number;
  /** CSS selectors excluding matching descendants and their entire subtrees.
   * Malformed selectors throw `SyntaxError` after mount, even when loading is false.
   * The `data-skeleton-ignore` presence marker also excludes a subtree; omission uses no extra selectors.
   */
  ignoreSelectors?: readonly string[];
  /** Direct HTML injection is unsupported; provide rendered children instead. */
  dangerouslySetInnerHTML?: never;
};
