import * as React from 'react';
import { useSx } from '../../hooks/useSx';
import { useStageContext } from './stage-context.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageBodyProps } from './stage.types';

/** Render scrollable caller-owned content in the available Stage space.
 * Requires an owning Stage and inherits its surface padding; JSX remains unchanged.
 * @param props Native div attributes, forwarded ref and optional content.
 * @returns A growing zero-padding div; raw HTML injection is unsupported.
 * @example
 * ```tsx
 * <Stage presentation="dialog"><Stage.Body><input aria-label="Name" /></Stage.Body></Stage>
 * ```
 */
export const StageBody = React.forwardRef<HTMLDivElement, StageBodyProps>(({ children, className, ...nativeProps }, ref) => {
  const attributes: React.HTMLAttributes<HTMLDivElement> = nativeProps;
  delete attributes.dangerouslySetInnerHTML;
  useStageContext();
  const styles = useStageStyles(), sx = useSx();
  return <div {...attributes} ref={ref} data-stage-body="" className={sx(styles.body, className)}>{children}</div>;
});
StageBody.displayName = 'Stage.Body';
