import * as React from 'react';
import { Card } from '@fluentui/react-components';
import { GroupHeader, GroupText, hasGroupContent } from '../group-header.internal';
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

/** Render one settings row without cloning, labeling or managing its controls.
 * @param props Native row attributes, optional text/icon/action and caller-owned controls.
 * @returns A flex row whose controls wrap only when the available width requires it.
 * @example <SettingsGroup.Item title="Email"><Switch aria-label="Email alerts" /></SettingsGroup.Item>
 */
const SettingsGroupItem = React.forwardRef<HTMLDivElement, SettingsGroupItemProps>(({
  title, description, icon, action, children, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useSettingsGroupStyles();
  const titlePresent = hasGroupContent(title), descriptionPresent = hasGroupContent(description);
  const iconPresent = hasGroupContent(icon);
  return <div {...nativeProps} ref={ref} data-settings-group-item=""
    className={sx(styles.item, minWidth.zero, className)}>
    <div data-settings-group-row="" className={sx(flex.row, flex.wrap, alignItems.center, gap.medium, styles.row, minWidth.zero, minHeight.px(44))}>
      {(titlePresent || descriptionPresent || iconPresent) && <div data-settings-group-leading="" className={sx(styles.leading, flex.row, alignItems.center, gap.medium, flexItem.grow, flexItem.shrink)}>
        {iconPresent && <span data-group-item-icon="" aria-hidden="true" className={sx(styles.icon, flexItem.noShrink, foreground.subtle)}>{icon}</span>}
        {(titlePresent || descriptionPresent) && <div data-group-text="" className={sx(flex.column, minWidth.zero, gap.none)}>
          {titlePresent && <GroupText>{title}</GroupText>}
          {descriptionPresent && <GroupText description>{description}</GroupText>}
        </div>}
      </div>}
      {hasGroupContent(children) && <div data-settings-group-controls="" className={sx(styles.controls, flex.row, flex.wrap, alignItems.center, gap.small, flexItem.noShrink, minWidth.zero)}>{children}</div>}
      {hasGroupContent(action) && <div className={sx(styles.controls, flex.row, flex.wrap, alignItems.center, flexItem.noShrink, gap.small)}>{action}</div>}
    </div>
  </div>;
});
SettingsGroupItem.displayName = 'SettingsGroup.Item';

const SettingsGroupRoot = React.forwardRef<HTMLDivElement, SettingsGroupProps>(({
  title, description, icon, action, headingLevel = 2, appearance = 'filled', children, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useSettingsGroupStyles();
  return <div {...nativeProps} ref={ref} className={sx(styles.root, flex.column, gap.medium, width.full, minWidth.zero, container.inlineSize, className)}>
    <GroupHeader title={title} description={description} icon={icon} action={action} headingLevel={headingLevel} />
    <Card appearance={appearance} size="medium" focusMode="off" className={sx(styles.card, width.full, gap.none)}>{children}</Card>
  </div>;
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
 * @example <SettingsGroup title="Notifications"><SettingsGroup.Item title="Email"><Switch aria-label="Email alerts" /></SettingsGroup.Item></SettingsGroup>
 */
export const SettingsGroup = Object.assign(SettingsGroupRoot, { Item: SettingsGroupItem });
