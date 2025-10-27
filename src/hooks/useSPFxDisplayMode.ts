// useSPFxDisplayMode.ts
// Hook to access display mode (Read/Edit)

import { useAtomValue } from 'jotai';
import { DisplayMode } from '@microsoft/sp-core-library';
import { useSPFxContext } from './useSPFxContext';
import { spfxAtoms } from './../core/atoms';

/**
 * Return type for useSPFxDisplayMode hook
 */
export interface SPFxDisplayModeInfo {
  /** Current display mode (Read/Edit) */
  readonly mode: DisplayMode;
  
  /** Whether currently in Edit mode */
  readonly isEdit: boolean;
  
  /** Whether currently in Read mode */
  readonly isRead: boolean;
}

/**
 * Hook to access SPFx display mode (readonly)
 * 
 * Display mode controls whether the WebPart/Extension is in:
 * - Read mode (DisplayMode.Read): Normal viewing mode
 * - Edit mode (DisplayMode.Edit): Editing/configuration mode
 * 
 * Note: displayMode is readonly in SPFx and controlled by SharePoint.
 * It changes when the user clicks the Edit button in the page.
 * 
 * Useful for:
 * - Showing/hiding edit controls
 * - Conditional rendering based on mode
 * - Different layouts for read vs edit
 * 
 * @returns Display mode information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { mode, isEdit } = useSPFxDisplayMode();
 *   
 *   return (
 *     <div>
 *       <p>Mode: {isEdit ? 'Editing' : 'Reading'}</p>
 *       {isEdit && <EditControls />}
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxDisplayMode(): SPFxDisplayModeInfo {
  const { instanceId } = useSPFxContext();
  
  // Get display mode atom for this instance
  const displayModeAtom = spfxAtoms.displayMode(instanceId);
  
  // Read current mode (readonly)
  const modeValue = useAtomValue(displayModeAtom);
  
  // Default to Read mode if not set
  const mode = modeValue ?? DisplayMode.Read;
  
  const isEdit = mode === DisplayMode.Edit;
  const isRead = mode === DisplayMode.Read;
  
  return {
    mode,
    isEdit,
    isRead,
  };
}

/**
 * Hook to check if currently in Edit mode
 * Shortcut for useSPFxDisplayMode().isEdit
 * 
 * @returns true if in Edit mode, false otherwise
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const isEdit = useSPFxIsEdit();
 *   
 *   return isEdit ? <EditView /> : <ReadView />;
 * }
 * ```
 */
export function useSPFxIsEdit(): boolean {
  const { isEdit } = useSPFxDisplayMode();
  return isEdit;
}
