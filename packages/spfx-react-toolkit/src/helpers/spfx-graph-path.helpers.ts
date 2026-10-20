/**
 * Photo endpoint inputs. `userId` takes precedence over email; omitting both targets the current user. This helper constructs a path and does not fetch a photo.
 *
 * @example
 * ```ts
 * const path = buildUserPhotoEndpoint({ email: 'user@contoso.com', size: '48x48' });
 * ```
 */
type SPFxUserPhotoOptionsResult = {
  /**
   * Optional Graph user identifier; takes precedence over email when nonempty.
   */
  readonly userId?: string;
  /**
   * Optional user email/UPN when no nonempty userId is supplied. Omitted targets /me.
   */
  readonly email?: string;
  /**
   * Requested Graph photo size. Defaults to 240x240 when omitted.
   */
  readonly size?: '48x48' | '64x64' | '96x96' | '120x120' | '240x240' | '360x360' | '432x432' | '504x504' | '648x648';
};

/**
 * Builds the Graph content path under the current user app root. Folder characters outside letters, digits, hyphens and underscores become hyphens. The file name is interpolated unchanged; supply a path-safe name. This helper performs no I/O or folder creation.
 *
 * @param fileName - File name interpolated into the Graph path without encoding.
 * @param folder - Optional namespace folder; omitted or empty uses the app root directly.
 * @returns Graph path ending in :/content.
 *
 * @example
 * ```ts
 * const path = buildOneDriveAppDataPath('settings.json', 'dashboard');
 * ```
 */
export function buildOneDriveAppDataPath(fileName: string, folder?: string): string {
  const basePath = '/me/drive/special/approot:';

  if (folder) {
    const safeFolderName = folder.replace(/[^a-zA-Z0-9-_]/g, '-');
    return `${basePath}/${safeFolderName}/${fileName}:/content`;
  }

  return `${basePath}/${fileName}:/content`;
}

/**
 * Builds a Graph photo value endpoint. A nonempty userId takes precedence over email; omission targets /me. Identifiers are interpolated unchanged; this helper does not encode inputs or fetch image bytes.
 *
 * @param options - Optional user selector and size; size defaults to 240x240.
 * @returns Graph path ending in /photos/<size>/$value.
 *
 * @example
 * ```ts
 * const path = buildUserPhotoEndpoint({ email: 'user@contoso.com', size: '48x48' });
 * ```
 */
export function buildUserPhotoEndpoint(options?: SPFxUserPhotoOptionsResult): string {
  const {
    userId,
    email,
    size = '240x240',
  } = options || {};

  let basePath: string;

  if (userId) {
    basePath = `/users/${userId}`;
  } else if (email) {
    basePath = `/users/${email}`;
  } else {
    basePath = '/me';
  }

  return `${basePath}/photos/${size}/$value`;
}
