import { useMemo } from 'react';
import { makeStyles as createStyles } from '@griffel/core';
import { makeStyles } from '@griffel/react';
import { useSxStyleEnvironment } from '../../helpers/styles/renderer.internal';
import type { AutoSkeletonBone } from './auto-skeleton-geometry.internal';

/** Preserve mounted content layout while the sibling overlay represents loading. */
export const useAutoSkeletonStyles = makeStyles({
  root: {
    position: 'relative',
  },
  contentHidden: {
    // Opacity hides the whole subtree, including descendants with visibility: visible.
    opacity: 0,
    pointerEvents: 'none',
  },
  overlay: {
    position: 'absolute',
    inset: 0,
    overflow: 'hidden',
    pointerEvents: 'none',
  },
  item: {
    position: 'absolute',
  },
  staticItem: {
    // Keep Fluent wave's themed base and forced-colors fill, without its moving gradient.
    '::after': {
      animationName: 'none',
      backgroundImage: 'none',
      transform: 'none',
    },
  },
});

/** Compile current geometry with the host renderer; Griffel deduplicates repeated rules. */
export function useAutoSkeletonGeometryClasses(bones: readonly AutoSkeletonBone[]): readonly string[] {
  const { renderer, dir } = useSxStyleEnvironment();
  return useMemo(() => bones.map(bone => createStyles({
    root: {
      position: 'absolute',
      // Preserve physical measured coordinates without mixing Griffel class directions.
      left: `${bone.x}px /* @noflip */`,
      top: `${bone.y}px`,
      width: `${bone.width}px`,
      height: `${bone.height}px`,
      // A measured zero is content geometry, not an instruction to flatten Fluent's rectangle shape.
      ...(bone.shape === 'circle' ? { borderRadius: '50%' }
        : bone.radius > 0 ? { borderRadius: `${bone.radius}px` } : {}),
    },
  })({ renderer, dir }).root), [bones, renderer, dir]);
}
