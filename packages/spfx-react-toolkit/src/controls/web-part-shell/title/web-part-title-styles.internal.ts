import { makeStyles, type GriffelStyle } from '@griffel/react';
import { tokens, typographyStyles } from '@fluentui/react-theme';
import { useSx } from '../../../hooks/useSx';
import { body1 } from '../../../helpers/styles/typography';
import { primary } from '../../../helpers/styles/foreground';
import { titleFormatRegistry, getTitleLineClass, getTitleSizeClass } from './web-part-title-format.internal';

/** Shared scoped wrapper for Read blocks and the private editor's document DOM. */
export interface WebPartTitleStyleClasses {
  /** Apply to an ancestor of the document block; semantic markers never stand alone. */
  root: string;
}

/** Build finite protocol selectors only; no renderer or CSS insertion at module evaluation. */
function createTitleWrapperStyle(): GriffelStyle {
  const style: GriffelStyle = {
    minInlineSize: 0,
    overflowWrap: 'anywhere',
    whiteSpace: 'pre-wrap',
    '& h1, & h2, & h3, & h4, & p': {
      marginBlockStart: 0,
      marginBlockEnd: 0,
    },
    '& p': { ...typographyStyles[titleFormatRegistry.paragraphPreset.typography] },
    '& strong': { fontWeight: tokens.fontWeightBold },
    '& em': { fontStyle: 'italic' },
    '& u': { textDecorationLine: 'underline' },
    '& .wps-space-before': { marginBlockStart: tokens[titleFormatRegistry.spaceToken] },
    '& .wps-space-after': { marginBlockEnd: tokens[titleFormatRegistry.spaceToken] },
  };
  for (const level of [1, 2, 3, 4] as const) {
    style[`& h${level}`] = { ...typographyStyles[titleFormatRegistry.headingPresets[level].typography] };
  }
  for (let size = titleFormatRegistry.minSize; size <= titleFormatRegistry.maxSize; size++) {
    style[`& .${getTitleSizeClass(size)}`] = { fontSize: `${size}px` };
  }
  for (const line of titleFormatRegistry.lineSpacing) {
    style[`& .${getTitleLineClass(line)}`] = { lineHeight: line };
  }
  for (const color of ['primary', 'subtle', 'brand'] as const) {
    style[`& .wps-color-${color}`] = { color: tokens[titleFormatRegistry.colors[color]] };
  }
  return style;
}

// Griffel allocates its local style factory here, but uses the inherited renderer only in the hook.
// This is deliberately not annotated PURE; it does not create a renderer or insert global CSS.
const useScopedTitleStyles = makeStyles({ root: createTitleWrapperStyle() });

/** Reuse host Fluent typography, palette, direction and Griffel renderer in either title mode. */
export function useWebPartTitleStyles(): WebPartTitleStyleClasses {
  const sx = useSx();
  const styles = useScopedTitleStyles();
  return { root: sx(body1, primary, styles.root) };
}
