import * as React from 'react';
import { renderContentSlot } from '../content-slot.internal';
import { useSx } from '../../hooks/useSx';
import * as margin from '../../helpers/styles/margin';
import * as minWidth from '../../helpers/styles/min-width';
import * as gap from '../../helpers/styles/gap';
import * as typography from '../../helpers/styles/typography';
import * as foreground from '../../helpers/styles/foreground';
import { useDescriptionListStyles } from './description-list.styles.internal';
import type { DescriptionListLayout, DescriptionListProps, DescriptionListItemProps } from './description-list.types';

const DescriptionListContext = React.createContext<DescriptionListLayout | undefined>(undefined);

/** One pair with preserved native semantic cells and caller-owned content/state.
 * @param props Native grouping attributes and Fluent content slots.
 * @returns A grouping div containing one dt and one dd.
 * @example
 * ```tsx
 * <DescriptionList><DescriptionList.Item label="Owner">Ada</DescriptionList.Item></DescriptionList>
 * ```
 */
const DescriptionListItem = React.forwardRef<HTMLDivElement, DescriptionListItemProps>(({
  label, children, className, ...nativeProps
}, ref) => {
  const layout = React.useContext(DescriptionListContext);
  const sx = useSx(), styles = useDescriptionListStyles();
  if (layout === undefined) throw new Error('DescriptionList.Item must be rendered inside DescriptionList.');
  delete nativeProps.dangerouslySetInnerHTML;
  const itemLayout = layout === 'horizontal' ? styles.itemHorizontal : layout === 'vertical' ? styles.itemVertical : styles.itemAuto;
  return <div {...nativeProps} ref={ref} data-description-list-item="" className={sx(styles.item, itemLayout, minWidth.zero, className)}>
    <dt className={sx(styles.cell, margin.none, minWidth.zero)}>
      {renderContentSlot(label, 'div', sx, { shorthand: 'wrap', className: sx(typography.body1Strong, foreground.primary, minWidth.zero) })}
    </dt>
    <dd className={sx(styles.cell, margin.none, minWidth.zero)}>
      {renderContentSlot(children, 'div', sx, { shorthand: 'wrap', className: sx(typography.body1, foreground.primary, minWidth.zero) })}
    </dd>
  </div>;
});
DescriptionListItem.displayName = 'DescriptionList.Item';

const DescriptionListRoot = React.forwardRef<HTMLDListElement, DescriptionListProps>(({
  items, children, layout = 'auto', className, ...nativeProps
}, ref) => {
  const sx = useSx(), styles = useDescriptionListStyles();
  delete nativeProps.dangerouslySetInnerHTML;
  // React.ReactNode retains legacy null children; treat them like undefined without introducing null defaults.
  if (items !== undefined ? items.length === 0 : children === undefined || children === null || typeof children === 'boolean') return <></>;
  return <DescriptionListContext.Provider value={layout}>
    <dl {...nativeProps} ref={ref} data-description-list-layout={layout}
      className={sx(styles.root, margin.none, minWidth.zero, gap.medium, styles[layout], className)}>
      {items === undefined ? children : items.map(item => <DescriptionListItem key={item.id} label={item.label}>{item.value}</DescriptionListItem>)}
    </dl>
  </DescriptionListContext.Provider>;
});
DescriptionListRoot.displayName = 'DescriptionList';

/** Display read-only metadata for one record as a native description list.
 * Supply immutable items with stable IDs or compound Item children. Strings remain literal;
 * formatting, missing-value labels, loading and empty states belong to the consumer.
 * Auto layout stacks without a named container and shares two columns from 480px of the
 * nearest apvee-sx ancestor. Explicit horizontal/vertical overrides the query.
 * Uses Fluent theme variables and the host Griffel renderer, with no SPFx provider requirement.
 * No generated inline styles, keyboard behavior, cloning or application state.
 * Empty data/absent children render nothing; empty fragments can leave an empty dl.
 * @param props Native dl attributes, layout and one content pathway.
 * @returns A native dl of term/definition pairs, or an empty fragment.
 * @example
 * ```tsx
 * <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada Lovelace' }]} />
 * ```
 */
export const DescriptionList = Object.assign(DescriptionListRoot, {
  /** Render one term/value pair inside DescriptionList, preserving native dt/dd cells.
   * Native attributes and ref target the grouping div; Fluent slots customize cell content.
   * @param props Native grouping attributes, required label and associated value slot.
   * @returns A grouping div containing one dt and one dd; requires a DescriptionList parent.
   * @example
   * ```tsx
   * <DescriptionList><DescriptionList.Item label="Owner">Ada</DescriptionList.Item></DescriptionList>
   * ```
   */
  Item: DescriptionListItem,
});
