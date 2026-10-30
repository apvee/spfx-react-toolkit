import type * as React from 'react';
import type { StateMessagePreset } from '../state-message.types';
import type { ThemeMode } from '../state-message-theme.internal';
import { DefaultIllustration } from './default-illustration';
import { EmptyIllustration } from './empty-illustration';
import { ErrorIllustration } from './error-illustration';
import { HelloIllustration } from './hello-illustration';
import { ThanksIllustration } from './thanks-illustration';

/** Theme selection shared by the private decorative artwork. */
export interface IllustrationProps {
  readonly mode: ThemeMode;
}

const illustrations: Record<StateMessagePreset, React.ComponentType<IllustrationProps>> = {
  generic: DefaultIllustration,
  empty: EmptyIllustration,
  error: ErrorIllustration,
  success: ThanksIllustration,
  welcome: HelloIllustration
};

/** Resolve semantic presets without exposing the artwork as a public API. */
export function resolvePresetIllustration(preset: StateMessagePreset): React.ComponentType<IllustrationProps> {
  return illustrations[preset];
}
