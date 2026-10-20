import * as React from 'react';
import type { ContentTypePickerProps, ContentTypePickerOption } from './picker.types';
import type { ContentTypePickerValue } from './picker-values.types';
import { createContentTypePickerSource } from './content-type-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/** Select read-only SharePoint content type references from a collection root or one list.
 * Requires an SPFx toolkit provider with SPHttpClient. Site scope (default) resolves the owning
 * collection root even for a subweb URL; list scope preserves the owning web and exact list copy ID.
 * Defaults to Combobox, single selection and includeHidden=true. Opening retrieves every REST
 * page, sorts by Name then ID, and caches the complete catalog. Typed search matches name prefixes
 * locally and case-insensitively; no personal history or initial 30-item cap applies.
 * Inherits PickerProps model seeds, reactive availableOptions, array selection events, errors,
 * accessibility and fixed presentation slots. ReadOnly/Sealed do not remove choices. Invalid
 * IDs/configuration throw synchronously; request failures use the inherited onError contract.
 * @param props Catalog scope/list selector, optional group/ID/hidden filters, selection models and native presentation.
 * @returns A Fluent Field with Combobox or TagPicker. Single and multiple selections both use scoped reference arrays.
 * @example <ContentTypePicker label="Site content type" webUrl="https://tenant.sharepoint.com/sites/hr/sub" />
 * @example <ContentTypePicker label="List content type" scope="list" list="Documents" contentTypeId="0x0101" includeDescendants multiselect />
 */
export function ContentTypePicker(props: ContentTypePickerProps): React.ReactElement {
  const context = usePickerSourceContext();
  const source = createContentTypePickerSource(props, context);
  const runtime = usePickerRuntime<ContentTypePickerValue, ContentTypePickerOption>(props, source, { multiselect: false });
  return <PickerRenderer<ContentTypePickerValue, ContentTypePickerOption> props={props} runtime={runtime} defaultPresentation="combobox" defaultMultiselect={false} />;
}
