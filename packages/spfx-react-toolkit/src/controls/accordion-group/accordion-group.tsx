import * as React from 'react';
import { Accordion, AccordionItem, AccordionHeader, AccordionPanel, Card, useId } from '@fluentui/react-components';
import { GroupHeader, GroupText, hasGroupContent } from '../group-header.internal';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as gap from '../../helpers/styles/gap';
import * as width from '../../helpers/styles/width';
import * as minWidth from '../../helpers/styles/min-width';
import * as foreground from '../../helpers/styles/foreground';
import { useAccordionGroupStyles } from './accordion-group.styles.internal';
import type { AccordionGroupProps, AccordionGroupItemProps, AccordionGroupValue } from './accordion-group.types';

const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;
// React 17's HTML types omit native inert. An empty attribute enables inert; omission removes it.
type InertPanelProps = React.HTMLAttributes<HTMLDivElement> & { inert?: '' };

/** Render a Fluent disclosure, retaining inert descendants unless reset is requested.
 * @param props Native item attributes, value, trigger content and caller-owned panel children.
 * @returns A native Fluent AccordionItem with semantic heading, button and linked panel.
 * @example <AccordionGroup.Item value="profile" title="Profile">Details</AccordionGroup.Item>
 */
const AccordionGroupItem = React.forwardRef<HTMLDivElement, AccordionGroupItemProps>(({
  value, title, description, icon, headingLevel = 3, disabled = false, unmountOnClose = false, children, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useAccordionGroupStyles();
  const panelId = useId('accordion-group-panel');
  const titlePresent = hasGroupContent(title), descriptionPresent = hasGroupContent(description);
  const renderedIcon = hasGroupContent(icon)
    ? <span data-group-item-icon="" aria-hidden="true" className={sx(styles.icon, flexItem.noShrink, !disabled && foreground.subtle)}>{icon}</span>
    : undefined;
  return <AccordionItem {...nativeProps} ref={ref} value={value} disabled={disabled} data-accordion-group-item=""
    className={sx(styles.item, minWidth.zero, className)}>
    <AccordionHeader as={headings[headingLevel - 1]} button={{ 'aria-controls': panelId, className: styles.button }} icon={renderedIcon} expandIconPosition="end">
      {(titlePresent || descriptionPresent) && <span data-group-text="" className={sx(flex.column, flexItem.grow, minWidth.zero, gap.none)}>
        {titlePresent && <GroupText disabled={disabled}>{title}</GroupText>}
        {descriptionPresent && <GroupText description disabled={disabled}>{description}</GroupText>}
      </span>}
    </AccordionHeader>
    <AccordionPanel id={panelId} className={styles.panel} collapseMotion={{
      children: (Motion, motionProps) => {
        // The installed Fluent motion exposes visibility but does not inert retained panels.
        // Adapt its existing panel element, preserving refs, keys and lifecycle. No duplicated open state.
        const panel = React.isValidElement<InertPanelProps>(motionProps.children)
          ? React.cloneElement(motionProps.children, { inert: motionProps.visible ? undefined : '' })
          : motionProps.children;
        return <Motion {...motionProps} unmountOnExit={unmountOnClose}>{panel}</Motion>;
      },
    }}>
      <div data-accordion-group-panel-content="" className={sx(styles.panelContent, minWidth.zero)}>{children}</div>
    </AccordionPanel>
  </AccordionItem>;
});
AccordionGroupItem.displayName = 'AccordionGroup.Item';

const AccordionGroupRoot = React.forwardRef<HTMLDivElement, AccordionGroupProps>(({
  title, description, icon, action, headingLevel = 2, appearance = 'filled', children,
  openItems, defaultOpenItems, onToggle, multiple = false, collapsible = true, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx(), styles = useAccordionGroupStyles();
  return <div {...nativeProps} ref={ref} className={sx(styles.root, flex.column, gap.medium, width.full, minWidth.zero, className)}>
    <GroupHeader title={title} description={description} icon={icon} action={action} headingLevel={headingLevel} />
    <Card appearance={appearance} size="medium" focusMode="off" className={sx(styles.card, width.full, gap.none)}>
      <Accordion<AccordionGroupValue> openItems={openItems} defaultOpenItems={defaultOpenItems} onToggle={onToggle}
        multiple={multiple} collapsible={collapsible} className={sx(styles.accordion, minWidth.zero)}>{children}</Accordion>
    </Card>
  </div>;
});
AccordionGroupRoot.displayName = 'AccordionGroup';

/** Group native disclosures in a medium Fluent Card with an optional outside header.
 * Fluent owns expansion, disabled behavior and motion; no SPFx provider is required.
 * Closed descendants remain mounted and inert by default; unmountOnClose resets them after exit.
 * Strings remain literal. All text slots inherit toolkit presets; JSX phrasing children
 * retain their props and explicit styling. Trigger content must contain no interactive descendants.
 * Supply nonempty title or description content to name the trigger; native item aria-label
 * labels the item div and does not name its button. Before programmatic closure,
 * consumers move focus from panel content when necessary. Root titles use the headingLevel wrapper,
 * including JSX; avoid nested headings. Native Card padding is the single lateral inset.
 * @param props Native div attributes, header and native Fluent expansion options.
 * @returns A div containing the optional header and composed Accordion inside a Card.
 * @example <AccordionGroup title="Preferences" defaultOpenItems={['profile']}><AccordionGroup.Item value="profile" title="Profile">Details</AccordionGroup.Item></AccordionGroup>
 */
export const AccordionGroup = Object.assign(AccordionGroupRoot, { Item: AccordionGroupItem });
