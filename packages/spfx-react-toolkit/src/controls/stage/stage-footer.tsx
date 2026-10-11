import * as React from 'react';
import { useSx } from '../../hooks/useSx';
import { useStageContext } from './stage-context.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageFooterProps } from './stage.types';

/** Render free caller-owned footer content without imposing actions or automatic dismissal.
 * Requires an owning Stage. Omit Footer entirely when no footer is wanted, leaving no gap.
 * @param props Native div attributes, forwarded ref and optional content.
 * @returns A nonshrinking zero-padding div preserving caller JSX and events.
 * @example <Stage.Footer><button type="submit">Save</button></Stage.Footer>
 */
export const StageFooter = React.forwardRef<HTMLDivElement, StageFooterProps>(({ children, className, ...nativeProps }, ref) => {
  const attributes: React.HTMLAttributes<HTMLDivElement> = nativeProps;
  delete attributes.dangerouslySetInnerHTML;
  useStageContext();
  const styles = useStageStyles(), sx = useSx();
  return <div {...attributes} ref={ref} data-stage-footer="" className={sx(styles.footer, className)}>{children}</div>;
});
StageFooter.displayName = 'Stage.Footer';
