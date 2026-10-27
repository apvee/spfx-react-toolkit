import * as resource from 'ApveeSpfxReactToolkitWebPartShellStrings';
import type { WebPartShellResource } from './web-part-shell-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativeWebPartShellResource: WebPartShellResource & { readonly __locale: string } = resource;
