import type { GriffelRenderer } from '@griffel/core';
import { useRenderer_unstable } from '@griffel/react';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import type { SxOptions } from './types';

export interface SxStyleEnvironment {
  readonly renderer: GriffelRenderer;
  readonly dir: 'ltr' | 'rtl';
}
/** Read existing host contexts exactly once per hook render; never create a renderer. */
export function useSxStyleEnvironment(options?: SxOptions): SxStyleEnvironment {
  const renderer = useRenderer_unstable();
  const { dir } = useFluent_unstable();
  return { renderer, dir: options?.dir ?? dir };
}
