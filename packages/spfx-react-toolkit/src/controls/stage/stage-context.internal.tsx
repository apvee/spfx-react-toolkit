import * as React from 'react';
import { useId } from '@fluentui/react-utilities';
import type { StageController } from './stage-controller.internal';

/** Per-root composition data; never shared between Stage instances. */
export interface StageContextValue {
  controller: StageController;
  surfaceId: string;
  titleId: string;
  descriptionId: string;
  hasDescription: boolean;
  registerDescription(present: boolean): () => void;
}

export const StageContext = React.createContext<StageContextValue | undefined>(undefined);

/** Require an owning Stage for each compound part. */
export function useStageContext(): StageContextValue {
  const context = React.useContext(StageContext);
  if (context === undefined) throw new Error('Stage parts must be rendered inside Stage.');
  return context;
}

/** Create stable native associations; caller id owns both surface and trigger aria-controls. */
export function useStageContextValue(controller: StageController, nativeSurfaceId?: string): StageContextValue {
  const surfaceId = useId('stage-surface-', nativeSurfaceId);
  const titleId = useId('stage-title-');
  const descriptionId = useId('stage-description-');
  const [descriptionCount, setDescriptionCount] = React.useState(0);
  const registerDescription = React.useCallback((present: boolean): (() => void) => {
    if (!present) return () => {};
    setDescriptionCount(count => count + 1);
    return () => setDescriptionCount(count => count - 1);
  }, []);
  return React.useMemo(() => ({ controller, surfaceId, titleId, descriptionId,
    hasDescription: descriptionCount > 0, registerDescription,
  }), [controller, surfaceId, titleId, descriptionId, descriptionCount, registerDescription]);
}
