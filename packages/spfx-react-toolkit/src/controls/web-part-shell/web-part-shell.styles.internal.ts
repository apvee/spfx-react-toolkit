import { makeStyles, useRenderer_unstable } from '@griffel/react';
import { makeStyles as makeCoreStyles } from '@griffel/core';
import { useSx } from '../../hooks/useSx';
import { inlineSize } from '../../helpers/styles/container';
import { zero as minWidthZero } from '../../helpers/styles/min-width';
import * as padding from '../../helpers/styles/padding';
import { medium as atMedium } from '../../helpers/styles/responsive';
import { column, row, wrap } from '../../helpers/styles/flex';
import { medium as mediumGap, small as smallGap } from '../../helpers/styles/gap';
import { center as centerItems } from '../../helpers/styles/align-items';
import type { WebPartShellPadding } from './web-part-shell.types';

const getProviderStyles = makeCoreStyles({
  provider: { backgroundColor: 'transparent', minWidth: '0px' },
});
const useStyles = makeStyles({
  title: { flexGrow: 1, flexShrink: 1, minWidth: '0px' },
  actions: { flexShrink: 0, minWidth: '0px' },
});

/** Class factories use the inherited renderer; module evaluation creates no renderer or CSS. */
export function useWebPartShellProviderStyles(dir: 'ltr' | 'rtl'): { provider: string } {
  const renderer = useRenderer_unstable();
  // Fluent merges its root classes using its effective dir, before providing that
  // direction to descendants. Match it explicitly without creating a renderer.
  return getProviderStyles({ renderer, dir });
}

/** Call below FluentProvider so direction and tokens match the shell theme. */
export function useWebPartShellLayoutStyles(surface: 'section' | 'application', requested: WebPartShellPadding): {
  root: string; content: string; header: string; title: string; actions: string; body: string;
} {
  const sx = useSx();
  const classes = useStyles();
  const automaticApplication = requested === 'auto' && surface === 'application';
  return {
    root: sx(inlineSize, minWidthZero),
    content: sx(column, mediumGap, minWidthZero, requested === 'auto'
      ? automaticApplication ? padding.medium : padding.none
      : padding[requested], automaticApplication && atMedium(padding.large)),
    header: sx(row, wrap, centerItems, smallGap, minWidthZero),
    title: sx(classes.title),
    actions: sx(classes.actions),
    body: sx(minWidthZero),
  };
}
