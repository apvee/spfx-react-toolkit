import * as resource from 'ApveeSpfxReactToolkitPickersStrings';
import type { PickersResource } from './pickers-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativePickersResource: PickersResource & { readonly __locale: string } = resource;
