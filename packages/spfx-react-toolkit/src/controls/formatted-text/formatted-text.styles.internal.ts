import { useMemo } from 'react';
import { makeStyles as createStyles } from '@griffel/core';
import { useSxStyleEnvironment } from '../../helpers/styles/renderer.internal';
import { makeStyles } from '@griffel/react';
import { tokens } from '@fluentui/react-theme';

/** Only control-specific declarations; common presentation uses the existing sx catalog. */
export const useFormattedTextStyles = makeStyles({
  root: { minInlineSize: 0, overflowWrap: 'break-word' },
  monospace: { fontFamily: tokens.fontFamilyMonospace },
  numeric: { fontFamily: tokens.fontFamilyNumeric },
  justify: { textAlign: 'justify' },
  italic: { fontStyle: 'italic' },
  underline: { textDecorationLine: 'underline' },
  strikethrough: { textDecorationLine: 'line-through' },
  combinedDecoration: { textDecorationLine: 'underline line-through' },
  clamp: { display: '-webkit-box', WebkitBoxOrient: 'vertical', overflow: 'hidden', textOverflow: 'ellipsis' },
  code: {
    fontFamily: tokens.fontFamilyMonospace, overflowWrap: 'anywhere',
    '@media (forced-colors: active)': {
      backgroundColor: 'Canvas', color: 'CanvasText', outline: '1px solid CanvasText',
    },
  },
  link: { overflowWrap: 'anywhere' },
  highlightedLink: { color: 'inherit', ':hover': { color: 'inherit' }, ':active': { color: 'inherit' } },
  mark: {
    backgroundColor: tokens.colorPaletteYellowBackground2,
    boxDecorationBreak: 'clone', WebkitBoxDecorationBreak: 'clone',
    '@media (forced-colors: active)': { backgroundColor: 'Highlight', color: 'HighlightText', forcedColorAdjust: 'none' },
  },
});

/** Compile a dynamic line count with the host renderer; repeated rules are deduplicated by Griffel. */
export function useLineClampClass(maxLines: number | undefined): string {
  const { renderer, dir } = useSxStyleEnvironment();
  return useMemo(() => maxLines === undefined ? ''
    : createStyles({ root: { WebkitLineClamp: maxLines } })({ renderer, dir }).root,
  [maxLines, renderer, dir]);
}
