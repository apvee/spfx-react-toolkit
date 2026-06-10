import type { MSGraphClientV3 } from '@microsoft/sp-http';

import { buildUserPhotoEndpoint } from '../helpers/spfx-graph-path.helpers';

export type SPFxUserPhotoServiceSize =
  | '48x48'
  | '64x64'
  | '96x96'
  | '120x120'
  | '240x240'
  | '360x360'
  | '432x432'
  | '504x504'
  | '648x648';

export interface SPFxUserPhotoServiceOptions {
  readonly userId?: string;
  readonly email?: string;
  readonly size?: SPFxUserPhotoServiceSize;
}

export interface SPFxUserPhotoService {
  getPhotoBlob: (options?: SPFxUserPhotoServiceOptions) => Promise<Blob>;
}

export function createSPFxUserPhotoService(
  graphClient: MSGraphClientV3
): SPFxUserPhotoService {
  const getPhotoBlob = async (options?: SPFxUserPhotoServiceOptions): Promise<Blob> => {
    const endpoint = buildUserPhotoEndpoint(options);
    const blob = await graphClient.api(endpoint).get();

    return blob as Blob;
  };

  return {
    getPhotoBlob
  };
}
