// atoms.ts
// Jotai atoms using atomFamily for per-instance scoping
// Each field has its own atomFamily for clean, direct access

import { atom } from 'jotai';
import { atomFamily } from 'jotai/utils';
import type { DisplayMode } from '@microsoft/sp-core-library';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { ContainerSize } from './types';

/**
 * Collection of atomFamily instances, one per state field
 * Each atomFamily creates scoped atoms per instanceId
 * 
 * Usage in hooks:
 *   const theme = useAtomValue(spfxAtoms.theme(instanceId));
 *   const setDisplayMode = useSetAtom(spfxAtoms.displayMode(instanceId));
 * 
 * This approach:
 * - Provides automatic scoping by instanceId
 * - Allows direct access to individual atoms
 * - Type-safe and follows Jotai best practices
 * - Memory efficient (Jotai handles cleanup)
 * 
 * Memory Management:
 * - Atoms are automatically cleaned up when SPFxProvider unmounts
 * - Each atomFamily.remove(instanceId) is called in cleanup effect
 * - This prevents memory leaks when SPFx components are disposed
 */
export const spfxAtoms = {
  // Core state atoms
  theme: atomFamily((_instanceId: string) => 
    atom<IReadonlyTheme | undefined>(undefined)
  ),
  
  displayMode: atomFamily((_instanceId: string) => 
    atom<DisplayMode | undefined>(undefined)
  ),
  
  properties: atomFamily((_instanceId: string) => 
    atom<unknown>(undefined)
  ),
  
  // Container state atoms
  containerEl: atomFamily((_instanceId: string) => 
    atom<HTMLElement | undefined>(undefined)
  ),
  
  containerSize: atomFamily((_instanceId: string) => 
    atom<ContainerSize | undefined>(undefined)
  ),
  
  // Teams context state (async initialized)
  teams: atomFamily((_instanceId: string) => 
    atom<{
      supported: boolean;
      context?: unknown;
      theme?: 'default' | 'dark' | 'highContrast';
      initialized: boolean;
    }>({ supported: false, initialized: false })
  ),
  
  // Dynamic data state (for cross-webpart communication)
  dynamicData: atomFamily((_instanceId: string) => 
    atom<{
      sources: Record<string, unknown>;
      initialized: boolean;
    }>({ sources: {}, initialized: false })
  ),
} as const;

/**
 * Helper type to get the atom type for a specific field
 * Useful for creating custom hooks that need atom references
 * 
 * Example:
 *   type ThemeAtom = SPFxAtomType<'theme'>;
 */
export type SPFxAtomType<K extends keyof typeof spfxAtoms> = 
  ReturnType<typeof spfxAtoms[K]>;
