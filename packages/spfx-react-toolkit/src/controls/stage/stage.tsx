import * as React from 'react';
import { StageHeader } from './stage-header';
import { StageBody } from './stage-body';
import { StageFooter } from './stage-footer';
import { StageDialog } from './stage-dialog.internal';
import { StageDrawer } from './stage-drawer.internal';
import { StagePopover } from './stage-popover.internal';
import { useStageController } from './stage-controller.internal';
import { StageContext, useStageContextValue } from './stage-context.internal';
import type { StageController } from './stage-controller.internal';
import type { StageProps, StageSurfaceAttributes } from './stage.types';

/** Keep the active anchor/placement when a different presentation is pending.
 * Common behavior remains live; the controller already owns state and compatible size.
 */
function useAdapterOptions(props: StageProps, controller: StageController): StageProps {
  const compatible = React.useRef(props);
  React.useLayoutEffect(() => {
    if (props.presentation === controller.presentation) compatible.current = props;
  }, [props, controller.presentation]);
  if (props.presentation === controller.presentation) return props;

  const previous = compatible.current;
  const state = props.open === undefined ? { defaultOpen: props.defaultOpen } : { open: props.open };
  const common = {
    ...state,
    children: props.children,
    onOpenChange: props.onOpenChange,
    dismissMode: props.dismissMode,
    unmountOnClose: props.unmountOnClose,
    mountNode: props.mountNode,
  };
  if (previous.presentation === 'popover') {
    return { ...common, presentation: 'popover', trigger: previous.trigger,
      positioning: previous.positioning, withArrow: previous.withArrow,
      size: controller.size === 'full' ? 'large' : controller.size };
  }
  if (previous.presentation === 'drawer') {
    return { ...common, presentation: 'drawer', trigger: previous.trigger,
      drawerPosition: previous.drawerPosition, size: controller.size };
  }
  return { ...common, presentation: 'dialog', trigger: previous.trigger, size: controller.size };
}

function getSurfaceAttributes(props: StageProps): StageSurfaceAttributes {
  const attributes: React.HTMLAttributes<HTMLDivElement> = { ...props };
  // Untyped consumers can still supply owned fields; never leak configuration,
  // raw HTML or a conflicting surface role/association into native markup.
  for (const key of ['presentation', 'size', 'trigger', 'drawerPosition', 'positioning', 'withArrow', 'open', 'defaultOpen',
    'onOpenChange', 'dismissMode', 'unmountOnClose', 'mountNode', 'children', 'title',
    'dangerouslySetInnerHTML', 'role', 'aria-modal', 'aria-labelledby', 'aria-describedby']) {
    Reflect.deleteProperty(attributes, key);
  }
  return attributes;
}

const StageRoot = React.forwardRef<HTMLDivElement, StageProps>(function StageRoot(props, ref) {
  const controller = useStageController(props);
  const options = useAdapterOptions(props, controller);
  const surfaceProps = getSurfaceAttributes(props);
  const context = useStageContextValue(controller, surfaceProps.id);
  const adapter = { controller, options, surfaceProps, surfaceRef: ref, children: props.children };
  return <StageContext.Provider value={context}>
    {controller.presentation === 'dialog' ? <StageDialog {...adapter} />
      : controller.presentation === 'drawer' ? <StageDrawer {...adapter} />
        : <StagePopover {...adapter} />}
  </StageContext.Provider>;
});
StageRoot.displayName = 'Stage';

/** Present one composed content structure in a Dialog, Drawer or anchored Popover.
 * Use Header, Body and optional Footer in declared order; a caller-owned form wrapper
 * may contain all three. Presentation is explicit and latched for each open session.
 * Strings in Header stay literal with toolkit presets; caller JSX remains unchanged.
 * Requires shared Fluent theme/renderer, but no SPFx provider. Controlled consumers
 * accept or refuse requests by updating open; native event prevention is not a callback veto.
 * @param props Surface attributes, presentation-specific options and open-state ownership.
 * @returns The integrated trigger, when supplied, and the Fluent-owned surface/portal.
 * @example
 * ```tsx
 * <Stage presentation="dialog" trigger={<button type="button">Edit</button>}>
 *   <Stage.Header title="Profile" closeLabel="Close" />
 *   <Stage.Body><input aria-label="Name" /></Stage.Body>
 *   <Stage.Footer><button type="button">Save</button></Stage.Footer>
 * </Stage>
 * ```
 */
export const Stage = Object.assign(StageRoot, {
  /** Semantic heading, decorative icon, supporting text and caller-owned action. Requires the owning Stage.
   * @param props Native div attributes and header content.
   * @returns The header div within the active surface; its ref targets that div.
   * @example
   * ```tsx
   * <Stage presentation="dialog"><Stage.Header title="Profile" /></Stage>
   * ```
   */
  Header: StageHeader,
  /** Growing scroll region, preserving the consumer's content. Requires the owning Stage.
   * @param props Native div attributes and body content.
   * @returns The body div within the active surface; its ref targets that div.
   * @example
   * ```tsx
   * <Stage presentation="dialog"><Stage.Body>Content</Stage.Body></Stage>
   * ```
   */
  Body: StageBody,
  /** Optional free-content footer; actions do not dismiss automatically. Requires the owning Stage.
   * @param props Native div attributes and footer content.
   * @returns The footer div within the active surface; its ref targets that div.
   * @example
   * ```tsx
   * <Stage presentation="dialog"><Stage.Footer><button type="button">Save</button></Stage.Footer></Stage>
   * ```
   */
  Footer: StageFooter,
});
