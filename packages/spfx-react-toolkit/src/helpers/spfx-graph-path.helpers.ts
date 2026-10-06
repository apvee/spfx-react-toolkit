type SPFxUserPhotoOptionsResult = {
  readonly userId?: string;
  readonly email?: string;
  readonly size?: '48x48' | '64x64' | '96x96' | '120x120' | '240x240' | '360x360' | '432x432' | '504x504' | '648x648';
};

/**
 * Builds a Microsoft Graph path for a file in the user's OneDrive app root.
 *
 * @param fileName - Name of the file
 * @param folder - Optional folder namespace
 * @returns Full Graph API path for file content
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
 * Builds the Microsoft Graph profile photo endpoint.
 *
 * @param options - Optional user identifier and photo size
 * @returns Graph endpoint for the requested user photo
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
