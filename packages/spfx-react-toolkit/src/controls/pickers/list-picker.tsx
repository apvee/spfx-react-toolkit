import * as React from 'react';
import type { ListPickerProps, ListPickerOption } from './picker.types';
import { createListPickerSource } from './list-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import type { ListPickerValue } from './picker-values.types';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/** Search SharePoint lists through the nearest SPFx provider; consumer candidates are independent of selection.
 * Defaults: Dropdown (historical combobox presentation), single selection, includeHidden=true, Title ascending, search from two characters.
 * Opening the menu loads up to 30 available options (initialCatalogLimit: integer 1–200). Typed search is case insensitive and independent of this catalog.
 * @param props Query source, scoped models, selection and Fluent presentation options.
 * @returns Accessible Fluent Field and picker. No persistence or provisioning.
 * @example <ListPicker label="Lista" baseTemplate={[100,101]} onOptionSelect={(_,data) => console.log(data.selectedValues)} />
 */
export function ListPicker(props: ListPickerProps): React.ReactElement {
  const context = usePickerSourceContext(), source = createListPickerSource(props, context);
  const runtime = usePickerRuntime<ListPickerValue, ListPickerOption>(props, source, { multiselect: false, initialCatalogLimit: props.initialCatalogLimit ?? 30 });
  return <PickerRenderer<ListPickerValue, ListPickerOption> props={props} runtime={runtime} defaultPresentation="combobox" defaultMultiselect={false} />;
}
