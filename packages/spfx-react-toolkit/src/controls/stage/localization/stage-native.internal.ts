import * as resource from 'ApveeSpfxReactToolkitStageStrings';
import type { StageResource } from './stage-resource.types.internal';

/** Owning family boundary; SPFx selects the catalog and its plural grammar metadata. */
export const nativeStageResource: StageResource & { readonly __locale: string } = resource;
