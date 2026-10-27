import * as React from 'react';
import { Button, Overflow, OverflowItem, Popover, PopoverSurface, PopoverTrigger, Tag, useOverflowMenu, useOverflowVisibility } from '@fluentui/react-components';
import type { ResolvedDataGridTag } from '../column-normalization.internal';
import type { DataGridStrings } from '../data-grid.types';
import { useDataGridCellStyles } from '../data-grid.styles.internal';
function tagKey(tag: ResolvedDataGridTag): string { return `${typeof tag.id}:${tag.id}`; }
function nativeTag(tag: ResolvedDataGridTag): React.ReactElement {
    const presentation = tag.presentation;
    const icon = presentation?.icon;
    return <Tag appearance={presentation?.appearance} shape={presentation?.shape} size={presentation?.size} icon={icon === undefined ? undefined : <span>{icon}</span>} title={tag.label}>{tag.label}</Tag>;
}
function TagOverflow({ tags, visible, strings }: {
    tags: readonly ResolvedDataGridTag[];
    visible: readonly ResolvedDataGridTag[];
    strings: DataGridStrings;
}): React.ReactElement {
    const { ref } = useOverflowMenu<HTMLButtonElement>();
    const { itemVisibility } = useOverflowVisibility();
    const styles = useDataGridCellStyles();
    const visibleIds = new Set(visible.map(tagKey));
    const hidden = tags.filter(tag => !visibleIds.has(tagKey(tag)) || itemVisibility[tagKey(tag)] === false);
    if (hidden.length === 0)
        return <></>;
    return <Popover>
    <PopoverTrigger disableButtonEnhancement><Button ref={ref} className={styles.tagOverflow} size="small" appearance="subtle" aria-label={strings.tagOverflow(hidden.length)}>+{hidden.length}</Button></PopoverTrigger>
    <PopoverSurface className={styles.tagList}>{hidden.map(tag => <React.Fragment key={tagKey(tag)}>{nativeTag(tag)}</React.Fragment>)}</PopoverSurface>
  </Popover>;
}
/** Native overflow measures actual tags; configured and responsive overflow remain recoverable. */
export function DataGridTagsCell({ tags, maxVisible, strings }: {
    tags: readonly ResolvedDataGridTag[];
    maxVisible: number | undefined;
    strings: DataGridStrings;
}): React.ReactElement {
    const styles = useDataGridCellStyles();
    if (tags.length === 0)
        return <span className={styles.text}>{strings.noTags}</span>;
    const visible = maxVisible === undefined ? tags : tags.slice(0, maxVisible);
    // A changed set/order gets a fresh native measurement; no hand-authored width estimates.
    const measurementKey = JSON.stringify(visible.map(tag => [tagKey(tag), tag.label, tag.presentation?.size, tag.presentation?.shape, tag.presentation?.appearance, tag.presentation?.icon !== undefined]));
    return <Overflow key={measurementKey} minimumVisible={0} hasHiddenItems={visible.length < tags.length}>
    <div className={styles.tags} data-grid-tags="">
      {visible.map(tag => <OverflowItem key={tagKey(tag)} id={tagKey(tag)}>{nativeTag(tag)}</OverflowItem>)}
      <TagOverflow tags={tags} visible={visible} strings={strings}/>
    </div>
  </Overflow>;
}
