import * as resource from 'ApveeSpfxReactToolkitDataGridStrings';
import type { DataGridResource } from './data-grid-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativeDataGridResource: DataGridResource & { readonly __locale: string } = resource;
