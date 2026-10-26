import { makeStyles, type GriffelStyle } from '@griffel/react';
import { tokens, typographyStyles } from '@fluentui/react-theme';
import { useSx } from '../../hooks/useSx';
import { body1 } from '../../helpers/styles/typography';
import { primary } from '../../helpers/styles/foreground';
import { editableTextColorDescriptors, getEditableTextColorCss, editableTextFormatRegistry, getEditableTextLineClass, getEditableTextSizeClass } from './editable-text-format.internal';

/** Shared scoped wrapper for Read blocks and the private editor's document DOM. */
export interface EditableTextStyleClasses {
  /** Apply to an ancestor of the document block; semantic markers never stand alone. */
  root: string;
}

/** Build finite protocol selectors only; no renderer or CSS insertion at module evaluation. */
function createEditableTextWrapperStyle(): GriffelStyle {
  const style: GriffelStyle = {
    minInlineSize: 0,
    overflowWrap: 'anywhere',
    whiteSpace: 'pre-wrap',
    '& h1, & h2, & h3, & h4, & p': {
      marginBlockStart: 0,
      marginBlockEnd: 0,
    },
    '& p': { ...typographyStyles[editableTextFormatRegistry.paragraphPreset.typography] },
    '& strong': { fontWeight: tokens.fontWeightBold },
    '& em': { fontStyle: 'italic' },
    '& u': { textDecorationLine: 'underline' },
    '& .wps-space-before': { marginBlockStart: tokens[editableTextFormatRegistry.spaceToken] },
    '& .wps-space-after': { marginBlockEnd: tokens[editableTextFormatRegistry.spaceToken] },
  };
  for (const level of [1, 2, 3, 4] as const) {
    style[`& h${level}`] = { ...typographyStyles[editableTextFormatRegistry.headingPresets[level].typography] };
  }
  for (let size = editableTextFormatRegistry.minSize; size <= editableTextFormatRegistry.maxSize; size++) {
    style[`& .${getEditableTextSizeClass(size)}`] = { fontSize: `${size}px` };
  }
  for (const line of editableTextFormatRegistry.lineSpacing) {
    style[`& .${getEditableTextLineClass(line)}`] = { lineHeight: line };
  }
  for (const { id } of editableTextColorDescriptors) {
    style[`& .wps-color-${id}`] = { color: getEditableTextColorCss(id) };
  }
  return style;
}

// Griffel allocates its local style factory here, but uses the inherited renderer only in the hook.
// This is deliberately not annotated PURE; it does not create a renderer or insert global CSS.
const useScopedEditableTextStyles = makeStyles({ root: createEditableTextWrapperStyle() });

/** Reuse host Fluent typography, palette, direction and Griffel renderer in either text mode. */
export function useEditableTextStyles(): EditableTextStyleClasses {
  const sx = useSx();
  const styles = useScopedEditableTextStyles();
  return { root: sx(body1, primary, styles.root) };
}
