import * as resource from 'ApveeSpfxReactToolkitPaginationStrings';
import type { PaginationResource } from './pagination-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativePaginationResource: PaginationResource & { readonly __locale: string } = resource;
