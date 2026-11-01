import * as React from 'react';
import type { ExecutableShellPort, ExecutableShellPresentation } from '../generated/jsdoc-examples/catalogue';
import { useShellDemoOptions } from '../WebPartShellDemoHost';

/** Only presentation crosses the adapter; unrelated host options remain owned by the host. */
export function useExecutableShellPort(): ExecutableShellPort {
  const { options, update } = useShellDemoOptions();
  const appearance = options.appearance ?? 'subtle';
  const size = options.size ?? 'medium';
  const titlePosition = options.titlePosition ?? 'outside';
  const padding = options.padding ?? 'auto';
  const updatePresentation = React.useCallback((patch: Partial<ExecutableShellPresentation>): void => {
    update(patch);
  }, [update]);
  return React.useMemo(() => ({ presentation: { appearance, size, titlePosition, padding }, updatePresentation }),
    [appearance, size, titlePosition, padding, updatePresentation]);
}
