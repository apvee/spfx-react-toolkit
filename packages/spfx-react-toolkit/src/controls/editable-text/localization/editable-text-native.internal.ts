import * as resource from 'ApveeSpfxReactToolkitEditableTextStrings';
import type { EditableTextResource } from './editable-text-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativeEditableTextResource: EditableTextResource & { readonly __locale: string } = resource;
