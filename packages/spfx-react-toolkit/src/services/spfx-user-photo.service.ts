import type { MSGraphClientV3 } from '@microsoft/sp-http';

import { buildUserPhotoEndpoint } from '../helpers/spfx-graph-path.helpers';

/**
 * Microsoft Graph photo size requested by the service; availability depends on the stored user photo.
 *
 * @example
 * ```ts
 * const size: SPFxUserPhotoServiceSize = '240x240';
 * ```
 */
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

/**
 * User and size selection for a Graph photo. If neither identifier is supplied, targets /me.
 *
 * @example
 * ```ts
 * const options: SPFxUserPhotoServiceOptions = { email: 'user@contoso.com', size: '96x96' };
 * ```
 */
export interface SPFxUserPhotoServiceOptions {
  /**
   * Graph user/object ID; takes precedence over email when nonempty.
   */
  readonly userId?: string;
  /**
   * User email or UPN, used when userId is absent/empty.
   */
  readonly email?: string;
  /**
   * Requested photo dimensions; defaults to 240x240.
   */
  readonly size?: SPFxUserPhotoServiceSize;
}

/**
 * Explicit profile-photo retrieval through a caller-owned Graph client. Does not create or revoke object URLs.
 *
 * @example
 * ```ts
 * const photos: SPFxUserPhotoService = createSPFxUserPhotoService(graphClient);
 * const blob = await photos.getPhotoBlob();
 * ```
 */
export interface SPFxUserPhotoService {
  /**
   * Fetches a profile photo.
   * @param options - Optional user identifier and photo dimensions.
   * @returns Graph response as a Blob.
   * @throws Graph failures, including missing photos and insufficient permissions.
   */
  getPhotoBlob: (options?: SPFxUserPhotoServiceOptions) => Promise<Blob>;
}

/**
 * Creates profile-photo operations without React state or provider access.
 * @param graphClient - Initialized Graph client with the required permissions.
 * @returns Photo retrieval methods; construction performs no request.
 *
 * @example
 * ```ts
 * const photos = createSPFxUserPhotoService(graphClient);
 * const blob = await photos.getPhotoBlob({ email: 'user@contoso.com' });
 * ```
 */
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
