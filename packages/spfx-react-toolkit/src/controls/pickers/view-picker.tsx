import * as React from 'react';
import type { ViewPickerProps, ViewPickerOption } from './picker.types';
import { createViewPickerSource } from './view-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import type { ViewPickerValue } from './picker-values.types';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/** Search public/personal SharePoint list views through an SPFx provider.
 * Defaults: Dropdown (historical combobox presentation), single selection, Title ascending; a list is required.
 * Opening the menu loads all matching views, following native paging. Typed search filters the complete memory-cached catalog locally, case-insensitively; no initial display limit.
 * @param props Exact toolkit list selector, filters, reference models and presentation.
 * @returns A native Fluent Field and view picker; selection values always remain scoped arrays.
 * @example <ViewPicker label="Vista" list={{kind:'id',id:'11111111-1111-4111-8111-111111111111'}} />
 */
export function ViewPicker(props: ViewPickerProps): React.ReactElement {
  const context = usePickerSourceContext(), source = createViewPickerSource(props, context);
  const runtime = usePickerRuntime<ViewPickerValue, ViewPickerOption>(props, source, { multiselect: false });
  return <PickerRenderer<ViewPickerValue, ViewPickerOption> props={props} runtime={runtime} defaultPresentation="combobox" defaultMultiselect={false} showBlankOption={props.showBlankOption} />;
}
