import * as React from 'react';
import { inspectContentSlot, renderContentSlot } from '../content-slot.internal';
import { useId } from '@fluentui/react-utilities';
import { SettingsFieldContext } from '../settings-field-context.internal';
import type { SettingsFieldTarget } from '../settings-field-context.internal';
import { Card, Label } from '@fluentui/react-components';
import { GroupHeader, GroupText, hasGroupContent, hasGroupDescriptionContent } from '../group-header.internal';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as gap from '../../helpers/styles/gap';
import * as width from '../../helpers/styles/width';
import * as minWidth from '../../helpers/styles/min-width';
import * as minHeight from '../../helpers/styles/min-height';
import * as foreground from '../../helpers/styles/foreground';
import * as container from '../../helpers/styles/container';
import { useSettingsGroupStyles } from './settings-group.styles.internal';
import type { SettingsGroupProps, SettingsGroupItemProps } from './settings-group.types';

/** Inspect only static phrasing children; custom components own their rendered title. */
function hasSettingsFieldTitle(value: unknown): boolean {
  const inspected = inspectContentSlot(value);
  if (inspected.kind === 'render') return true;
  if (inspected.kind === 'absent') return false;
  return hasStaticSettingsFieldTitle(inspected.children);
}

function hasStaticSettingsFieldTitle(content: React.ReactNode): boolean {
  return React.Children.toArray(content).some(child => {
    if (typeof child === 'string') return child.trim().length > 0;
    if (typeof child === 'number') return true;
    if (React.isValidElement<{ children?: React.ReactNode }>(child)) {
      return typeof child.type === 'string' || child.type === React.Fragment
        ? hasStaticSettingsFieldTitle(child.props.children)
        : true;
    }
    return false;
  });
}

/** Render one settings row; field mode labels one built-in Form logical control.
 * @param props Native row attributes, optional text/icon/action and caller-owned controls.
 * @returns A flex row whose controls wrap only when the available width requires it.
 * @example
 * ```tsx
 * <SettingsGroup><SettingsGroup.Item title="Email"><Switch aria-label="Email alerts" /></SettingsGroup.Item></SettingsGroup>
 * ```
 */
const SettingsGroupItem = React.forwardRef<HTMLDivElement, SettingsGroupItemProps>(({
  title, description, icon, action, children, className, field = false, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useSettingsGroupStyles();
  const titlePresent = hasGroupContent(title), descriptionPresent = hasGroupContent(description);
  const iconPresent = hasGroupContent(icon);
  const labelId = useId('settings-label-'), descriptionId = useId('settings-description-');
  const [target, setTarget] = React.useState<SettingsFieldTarget>();
  const activeOwner = React.useRef<symbol>();
  const register = React.useCallback((next: SettingsFieldTarget) => {
    if (activeOwner.current !== undefined && activeOwner.current !== next.owner) {
      throw new TypeError('SettingsGroup.Item field requires exactly one built-in Form logical control.');
    }
    activeOwner.current = next.owner;
    setTarget(next);
    return () => {
      if (activeOwner.current === next.owner) {
        activeOwner.current = undefined;
        setTarget(undefined);
      }
    };
  }, []);
  React.useLayoutEffect(() => {
    if (field && activeOwner.current === undefined) {
      throw new TypeError('SettingsGroup.Item field requires exactly one built-in Form logical control.');
    }
  });
  const bridge = React.useMemo(() => field ? { labelId, descriptionId: hasGroupDescriptionContent(description) ? descriptionId : undefined, register } : undefined,
    [field, labelId, descriptionId, description, register]);
  if (field && !hasSettingsFieldTitle(title)) {
    throw new TypeError('SettingsGroup.Item field requires a meaningful title.');
  }
  return <div {...nativeProps} ref={ref} data-settings-group-item=""
    className={sx(styles.item, minWidth.zero, className)}>
    <div data-settings-group-row="" className={sx(flex.row, flex.wrap, alignItems.center, gap.medium, styles.row, minWidth.zero, minHeight.px(44))}>
      {(titlePresent || descriptionPresent || iconPresent) && <div data-settings-group-leading="" className={sx(styles.leading, flex.row, alignItems.center, gap.medium, flexItem.grow, flexItem.shrink)}>
        {iconPresent && <span data-group-item-icon="" aria-hidden="true" className={sx(styles.icon, flexItem.noShrink, foreground.subtle)}>{renderContentSlot(icon, 'span', sx, { shorthand: 'preserve', className: styles.iconContent })}</span>}
        {(titlePresent || descriptionPresent) && <div data-group-text="" className={sx(flex.column, minWidth.zero, gap.none)}>
          {titlePresent && (field ? <Label id={labelId} data-settings-field-label="" data-required={target?.required || undefined}
            htmlFor={target?.nativeLabel ? target.id : undefined} required={target?.required} disabled={target?.disabled}
            aria-disabled={target?.disabled || undefined}><GroupText disabled={target?.disabled}>{title}</GroupText></Label> : <GroupText>{title}</GroupText>)}
          {descriptionPresent && (field ? <span id={descriptionId}><GroupText description>{description}</GroupText></span> : <GroupText description>{description}</GroupText>)}
        </div>}
      </div>}
      {hasGroupContent(children) && <div data-settings-group-controls="" className={sx(styles.controls, flex.row, flex.wrap, alignItems.center, gap.small, flexItem.noShrink, minWidth.zero)}><SettingsFieldContext.Provider value={bridge}>{children}</SettingsFieldContext.Provider></div>}
      {hasGroupContent(action) && <div className={sx(styles.controls, flex.row, flex.wrap, alignItems.center, flexItem.noShrink, gap.small)}><SettingsFieldContext.Provider value={undefined}>{renderContentSlot(action, 'div', sx, { shorthand: 'preserve' })}</SettingsFieldContext.Provider></div>}
    </div>
  </div>;
});
SettingsGroupItem.displayName = 'SettingsGroup.Item';

const SettingsGroupRoot = React.forwardRef<HTMLDivElement, SettingsGroupProps>(({
  title, description, icon, action, headingLevel = 2, appearance = 'filled', children, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useSettingsGroupStyles();
  return <SettingsFieldContext.Provider value={undefined}><div {...nativeProps} ref={ref} className={sx(styles.root, flex.column, gap.medium, width.full, minWidth.zero, container.inlineSize, className)}>
    <GroupHeader title={title} description={description} icon={icon} action={action} headingLevel={headingLevel} />
    <Card appearance={appearance === 'alternate' ? 'filled-alternative' : appearance} size="medium" focusMode="off" className={sx(styles.card, width.full, gap.none)}>{children}</Card>
  </div></SettingsFieldContext.Provider>;
});
SettingsGroupRoot.displayName = 'SettingsGroup';

/** Group visible settings in a native medium Fluent Card. The optional header sits outside it.
 * Text and controls share a fluid wrapping row: compact controls remain beside text
 * while larger controls move below when they do not fit. Leading text grows from a 160px preferred basis, bounded by the available width.
 * Requires Fluent theme variables; no SPFx provider, cloned controls or form state.
 * Text uses semantic wrappers and toolkit typography; strings remain literal.
 * JSX phrasing children inherit presets without cloning and may supply explicit styling.
 * @param props Native div attributes, optional header and caller-owned settings rows.
 * @returns A div containing the optional header and Card.
 * @example
 * ```tsx
 * <SettingsGroup title="Notifications"><SettingsGroup.Item title="Email"><Switch aria-label="Email alerts" /></SettingsGroup.Item></SettingsGroup>
 * ```
 */
export const SettingsGroup = Object.assign(SettingsGroupRoot, {
  /** Render one settings row; ordinary rows retain caller-owned labels. `field` integrates
   * one built-in Form control with the row title/hint inside its Form root.
   * @param props Native row attributes, optional text/icon/action and primary controls.
   * @returns A native div row preserving caller controls and their state.
   * @example
   * ```tsx
   * <SettingsGroup><SettingsGroup.Item title="Email"><input aria-label="Email" /></SettingsGroup.Item></SettingsGroup>
   * ```
   */
  Item: SettingsGroupItem
});
