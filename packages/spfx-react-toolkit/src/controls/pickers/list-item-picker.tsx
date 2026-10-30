import * as React from 'react';
import type { ListItemPickerProps, ListItemPickerOption } from './picker.types';
import type { ListItemPickerValue } from './picker-values.types';
import { createListItemPickerSource } from './list-item-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/** Search actual list items/files/folders beneath an SPFx provider, beyond consumer-provided menu candidates.
 * Defaults: TagPicker, multiple selection, Id key, startsWith search, folders excluded.
 * Text/calculated/nontext fields use the appropriate REST/CAML projection with bounded paging.
 * @param props Required list/textField, query options, scoped models and presentation.
 * @returns A native Fluent picker with minimal JSON references; source fields are not persisted.
 * @example <ListItemPicker label="Documenti" list="Documents" textField="FileLeafRef" presentation="tagPicker" />
 */
export function ListItemPicker(props: ListItemPickerProps): React.ReactElement {
  const context = usePickerSourceContext(), source = createListItemPickerSource(props, context);
  const runtime = usePickerRuntime<ListItemPickerValue, ListItemPickerOption>(props, source, { multiselect: true });
  return <PickerRenderer<ListItemPickerValue, ListItemPickerOption> props={props} runtime={runtime} defaultPresentation="tagPicker" defaultMultiselect />;
}
