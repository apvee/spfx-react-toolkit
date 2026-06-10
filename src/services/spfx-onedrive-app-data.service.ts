import type { MSGraphClientV3 } from '@microsoft/sp-http';

import { buildOneDriveAppDataPath } from '../helpers/spfx-graph-path.helpers';

export interface SPFxOneDriveAppDataReadResult<T = unknown> {
  readonly data: T | undefined;
  readonly isNotFound: boolean;
}

export interface SPFxOneDriveAppDataService {
  read: <T = unknown>(
    fileName: string,
    folder?: string
  ) => Promise<SPFxOneDriveAppDataReadResult<T>>;
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

export function createSPFxOneDriveAppDataService(
  graphClient: MSGraphClientV3
): SPFxOneDriveAppDataService {
  const read = async <T = unknown>(
    fileName: string,
    folder?: string
  ): Promise<SPFxOneDriveAppDataReadResult<T>> => {
    const apiPath = buildOneDriveAppDataPath(fileName, folder);
    let fileContent: unknown;

    try {
      fileContent = await graphClient.api(apiPath).get();
    } catch (err) {
      if (isNotFoundError(err)) {
        return {
          data: undefined,
          isNotFound: true
        };
      }

      throw err;
    }

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
