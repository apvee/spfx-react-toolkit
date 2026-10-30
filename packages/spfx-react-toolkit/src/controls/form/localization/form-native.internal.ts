import * as resource from 'ApveeSpfxReactToolkitFormStrings';
import type { FormResource } from './form-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativeFormResource: FormResource & { readonly __locale: string } = resource;
