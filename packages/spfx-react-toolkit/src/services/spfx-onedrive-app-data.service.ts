import type { MSGraphClientV3 } from '@microsoft/sp-http';

import { buildOneDriveAppDataPath } from '../helpers/spfx-graph-path.helpers';

/**
 * JSON read outcome. T describes expected data and does not validate it.
 *
 * @example
 * ```ts
 * const result: SPFxOneDriveAppDataReadResult<{ theme: string }> = await appData.read('settings.json');
 * if (!result.isNotFound) console.log(result.data?.theme);
 * ```
 * @template T - Expected application data shape; no runtime validation is performed.
 */
export interface SPFxOneDriveAppDataReadResult<T = unknown> {
  /**
   * Parsed JSON, or undefined for a recognized missing file.
   */
  readonly data: T | undefined;
  /**
   * True only for recognized Graph metadata not-found failures; failed downloads reject.
   */
  readonly isNotFound: boolean;
}

/**
 * JSON reads and writes in the authenticated user's OneDrive app root. Requires a caller-owned Graph client with appropriate permissions and native fetch for reads.
 *
 * @example
 * ```ts
 * const appData: SPFxOneDriveAppDataService = createSPFxOneDriveAppDataService(graphClient);
 * await appData.write('settings.json', { theme: 'dark' }, 'dashboard');
 * ```
 */
export interface SPFxOneDriveAppDataService {
  /**
   * Reads metadata through Graph, then downloads JSON from its preauthenticated
   * HTTPS URL using native fetch without authorization headers or cookies.
   * This avoids Graph's browser-incompatible authenticated /content redirect.
   *
   * @param fileName - File name, passed into the Graph path without sanitization.
   * @param folder - Optional namespace; characters other than letters, digits, hyphen and underscore become hyphens.
   * @returns Parsed data, or isNotFound with undefined data for a missing file.
   * @throws Graph, download, invalid metadata or JSON parsing errors.
   * @example
   * ```ts
   * const result = await appData.read<Settings>('settings.json', 'my-app');
   * ```
   */
  read: <T = unknown>(
    fileName: string,
    folder?: string
  ) => Promise<SPFxOneDriveAppDataReadResult<T>>;
  /**
   * Writes JSON through Graph; existing content is replaced.
   * @param fileName - File name; passed into the Graph path without sanitization.
   * @param content - JSON-serializable value; generic types do not validate serialization.
   * @param folder - Optional namespace; non-alphanumeric characters except hyphen/underscore become hyphens.
   * @returns Resolves after Graph accepts the write.
   * @throws Serialization and Graph failures.
   */
  write: <T = unknown>(
    fileName: string,
    content: T,
    folder?: string
  ) => Promise<void>;
}

interface GraphErrorShape {
  readonly statusCode?: number;
  readonly status?: number;
  readonly code?: string;
  readonly message?: string;
  readonly body?: {
    readonly error?: {
      readonly code?: string;
      readonly message?: string;
    };
  };
}

function isNotFoundError(err: unknown): boolean {
  const graphError = (typeof err === 'object' && err !== null)
    ? err as GraphErrorShape
    : {};

  if (graphError.statusCode === 404 || graphError.status === 404) {
    return true;
  }

  const code = graphError.code ?? graphError.body?.error?.code;
  if (code && /itemnotfound/i.test(code)) {
    return true;
  }

  const message = graphError.message ?? graphError.body?.error?.message;
  if (message && /(\b404\b|not found|itemnotfound)/i.test(message)) {
    return true;
  }

  return false;
}

function parseOneDriveAppData<T>(fileContent: unknown): T {
  if (typeof fileContent === 'string') {
    try {
      return JSON.parse(fileContent) as T;
    } catch (parseError) {
      throw new Error(
        `Failed to parse JSON: ${parseError instanceof Error ? parseError.message : 'Unknown error'}`
      );
    }
  }

  return fileContent as T;
}

function getDownloadUrl(driveItem: unknown): string {
  const downloadUrl = typeof driveItem === 'object' && driveItem !== null &&
    '@microsoft.graph.downloadUrl' in driveItem
    ? driveItem['@microsoft.graph.downloadUrl']
    : undefined;
  if (typeof downloadUrl !== 'string') {
    throw new Error('OneDrive metadata did not provide a download URL.');
  }
  let parsedUrl: URL;
  try {
    parsedUrl = new URL(downloadUrl);
  } catch {
    throw new Error('OneDrive download URL is invalid.');
  }
  if (parsedUrl.protocol !== 'https:') {
    throw new Error('OneDrive download URL must use HTTPS.');
  }
  return downloadUrl;
}

/**
 * Creates explicit app-root JSON operations without React context or state.
 * @param graphClient - Initialized, caller-owned Microsoft Graph client.
 * @returns JSON read and write methods; construction performs no request.
 *
 * @example
 * ```ts
 * const appData = createSPFxOneDriveAppDataService(graphClient);
 * const settings = await appData.read<{ theme: string }>('settings.json');
 * ```
 */
export function createSPFxOneDriveAppDataService(
  graphClient: MSGraphClientV3
): SPFxOneDriveAppDataService {
  const read = async <T = unknown>(
    fileName: string,
    folder?: string
  ): Promise<SPFxOneDriveAppDataReadResult<T>> => {
    const apiPath = buildOneDriveAppDataPath(fileName, folder);
    let driveItem: unknown;

    try {
      // Graph /content redirects with 302, which fails browser CORS preflight
      // when the SDK supplies Authorization. Fetch its preauthenticated URL
      // directly, without forwarding Graph authorization or browser cookies.
      // Read complete metadata: projection on this special-folder path can
      // retain only id/etag and omit the download annotation.
      driveItem = await graphClient
        .api(apiPath.replace(/:\/content$/, ''))
        .get();
    } catch (err) {
      if (isNotFoundError(err)) {
        return {
          data: undefined,
          isNotFound: true
        };
      }

      throw err;
    }

    // Only Graph metadata can confirm absence. A failed or expired download
    // URL must remain an error so createIfMissing cannot overwrite a file.
    const response = await fetch(getDownloadUrl(driveItem), { credentials: 'omit' });
    if (!response.ok) {
      throw new Error(`Failed to download OneDrive file: HTTP ${response.status}`);
    }
    const fileContent = await response.text();

    return {
      data: parseOneDriveAppData<T>(fileContent),
      isNotFound: false
    };
  };

  const write = async <T = unknown>(
    fileName: string,
    content: T,
    folder?: string
  ): Promise<void> => {
    const apiPath = buildOneDriveAppDataPath(fileName, folder);
    const jsonContent = JSON.stringify(content);

    await graphClient
      .api(apiPath)
      .header('Content-Type', 'application/json')
      .put(jsonContent);
  };

  return {
    read,
    write
  };
}
