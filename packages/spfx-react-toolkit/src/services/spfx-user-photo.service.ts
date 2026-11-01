/**
 * Retrieve selected Microsoft Graph profile-photo bytes through a caller-owned client.
 *
 * @packageDocumentation
 * @docPage services/user-photo
 */
import type { MSGraphClientV3 } from '@microsoft/sp-http';

import { buildUserPhotoEndpoint } from '../helpers/spfx-graph-path.helpers';

/**
 * Microsoft Graph photo size requested by the service; availability depends on the stored user photo.
 *
 * @example
 * ```ts
 * import { SPFxUserPhotoServiceSize } from '@apvee/spfx-react-toolkit';
 *
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
 * import { SPFxUserPhotoServiceOptions } from '@apvee/spfx-react-toolkit';
 *
 * const options: SPFxUserPhotoServiceOptions = { email: 'user@contoso.com', size: '96x96' };
 * ```
 */
export interface SPFxUserPhotoServiceOptions {
  /**
   * Graph user/object ID; takes precedence over email when nonempty. Omitted/empty uses email, then /me.
   */
  readonly userId?: string;
  /**
   * User email or UPN, used when userId is absent/empty. Omitted/empty targets /me when there is no userId.
   */
  readonly email?: string;
  /**
   * Requested photo dimensions; omitted or undefined uses 240x240.
   * @defaultValue '240x240'
   */
  readonly size?: SPFxUserPhotoServiceSize;
}

/**
 * Explicit profile-photo retrieval through a caller-owned Graph client. Does not create or revoke object URLs.
 *
 * @example
 * ```ts
 * import type { MSGraphClientV3 } from '@microsoft/sp-http';
 * import { SPFxUserPhotoService, createSPFxUserPhotoService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(graphClient: MSGraphClientV3) {
 *   const photos: SPFxUserPhotoService = createSPFxUserPhotoService(graphClient);
 *   const blob = await photos.getPhotoBlob();
 * }
 * ```
 */
export interface SPFxUserPhotoService {
  /**
   * Fetches a profile photo.
   * @param options - Omitted targets /me at 240x240; userId takes precedence over email. Identifiers are interpolated without encoding.
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
 * @example Basic usage
 * ```ts
 * import type { MSGraphClientV3 } from '@microsoft/sp-http';
 * import { createSPFxUserPhotoService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(graphClient: MSGraphClientV3) {
 *   const photos = createSPFxUserPhotoService(graphClient);
 *   const blob = await photos.getPhotoBlob({ email: 'user@contoso.com' });
 * }
 * ```
 * @useWhen You need profile-photo bytes through a caller-owned Graph client without React or object-URL lifetime management.
 *
 *
 * @example Read own and selected-user photo blobs
 * ```ts
 * import type { MSGraphClientV3 } from '@microsoft/sp-http';
 *
 * import { createSPFxUserPhotoService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(graphClient: MSGraphClientV3) {
 *   const photos = createSPFxUserPhotoService(graphClient);
 *   const ownPhoto = await photos.getPhotoBlob();
 *   const selected = await photos.getPhotoBlob({ email: 'user@contoso.com', size: '96x96' });
 * }
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
