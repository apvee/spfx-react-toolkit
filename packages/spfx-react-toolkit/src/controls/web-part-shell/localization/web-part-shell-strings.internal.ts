import type { WebPartShellStrings } from '../web-part-shell.types';
import { nativeWebPartShellResource } from './web-part-shell-native.internal';

/** Shell naming resolves independently of common editor controls. */
export function resolveWebPartShellStrings(overrides?: Partial<WebPartShellStrings>): WebPartShellStrings {
  return {
    titleEditorLabel: overrides?.titleEditorLabel ?? nativeWebPartShellResource.titleEditorLabel,
    titlePlaceholder: overrides?.titlePlaceholder ?? nativeWebPartShellResource.titlePlaceholder,
  };
}
