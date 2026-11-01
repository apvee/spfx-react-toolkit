/**
 * WebPart read/edit flags follow the provider display mode.
 *
 * @packageDocumentation
 * @docPage hooks/properties
 */
// useSPFxDisplayMode.ts
// Hook to access display mode (Read/Edit)

import { useMemo } from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { useSPFxRuntimeSelector } from './../core/state.internal';

/**
 * Provider display mode and exact Read/Edit flags; missing runtime mode is read-only.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxDisplayModeInfo, useSPFxDisplayMode } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const mode: SPFxDisplayModeInfo = useSPFxDisplayMode();
 *   console.log(mode.isEdit);
 *   return <span>{mode.isEdit ? 'Edit' : 'Read'}</span>;
 * }
 * ```
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
 * Reads the provider runtime display mode and derives exact Read/Edit flags. An absent mode defaults to DisplayMode.Read. The host controls this value; this hook does not switch the page into edit mode. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Current host mode and Read/Edit boolean flags.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxDisplayMode } from '@apvee/spfx-react-toolkit';
 *
 * function ModeLabel() {
 *   const { isEdit } = useSPFxDisplayMode();
 *   return <span>{isEdit ? 'Editing' : 'Reading'}</span>;
 * }
 * ```
 *
 * @useWhen WebPart UI must distinguish the current edit and read modes.
 */
export function useSPFxDisplayMode(): SPFxDisplayModeInfo {
  const modeValue = useSPFxRuntimeSelector(state => state.displayMode);
  
  // Default to Read mode if not set
  const mode = modeValue ?? DisplayMode.Read;
  
  const isEdit = mode === DisplayMode.Edit;
  const isRead = mode === DisplayMode.Read;
  
  return useMemo(() => ({
    mode,
    isEdit,
    isRead,
  }), [mode, isEdit, isRead]);
}

/**
 * Returns useSPFxDisplayMode().isEdit. Missing runtime mode defaults to Read and therefore returns false. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns True only when the provider mode is DisplayMode.Edit.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxIsEdit } from '@apvee/spfx-react-toolkit';
 *
 * function EditLabel() {
 *   const editing = useSPFxIsEdit();
 *   return <span>{editing ? 'Editing' : 'Reading'}</span>;
 * }
 * ```
 *
 * @useWhen a WebPart only needs an edit-mode boolean for conditional UI.
 */
export function useSPFxIsEdit(): boolean {
  const { isEdit } = useSPFxDisplayMode();
  return isEdit;
}
