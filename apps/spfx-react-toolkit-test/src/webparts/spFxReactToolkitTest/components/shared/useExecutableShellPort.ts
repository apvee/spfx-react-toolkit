import type { ExecutableShellPort } from '../generated/jsdoc-examples/catalogue';
import { useShellPresentation } from '../WebPartShellDemoHost';

/** Forward only the bounded actual-Shell policy; recipes stay in owning examples. */
export function useExecutableShellPort(): ExecutableShellPort {
  return useShellPresentation();
}
