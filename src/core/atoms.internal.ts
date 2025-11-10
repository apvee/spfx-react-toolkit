// atoms.ts
// Jotai atoms for SPFx state management
// Each Provider instance gets its own isolated store

import { atom } from 'jotai';
import type { DisplayMode } from '@microsoft/sp-core-library';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { ContainerSize } from './types';

/**
 * Collection of Jotai atoms for SPFx state
 * 
 * These atoms are isolated per Provider instance through separate stores.
 * Each SPFxProvider creates its own store, ensuring complete isolation
 * between multiple instances on the same page.
 * 
 * Usage in hooks:
 *   const theme = useAtomValue(spfxAtoms.theme);
 *   const setDisplayMode = useSetAtom(spfxAtoms.displayMode);
 * 
 * Benefits:
 * - Simple atom definitions (no scoping complexity)
 * - Automatic isolation via Provider store
 * - Automatic cleanup when Provider unmounts
 * - Type-safe and follows Jotai best practices
 * 
 * @internal
 */
export const spfxAtoms = {
  // Core state atoms
  theme: atom<IReadonlyTheme | undefined>(undefined),
  
  displayMode: atom<DisplayMode | undefined>(undefined),
  
  properties: atom<unknown>(undefined),
  
  // Container state atoms
  containerEl: atom<HTMLElement | undefined>(undefined),
  
  containerSize: atom<ContainerSize | undefined>(undefined),
  
  // Teams context state (async initialized)
  teams: atom<{
    supported: boolean;
    context?: unknown;
    theme?: 'default' | 'dark' | 'highContrast';
    initialized: boolean;
  }>({ supported: false, initialized: false }),
} as const;
