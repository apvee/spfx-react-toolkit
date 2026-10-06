import { extractWebUrl } from '@pnp/sp';
import type { SPFI } from '@pnp/sp';
import type { IList } from '@pnp/sp/lists';

interface ListTargetSnapshot {
  readonly kind: string | undefined;
  readonly value: string | undefined;
}

const errorPrefix = '[createSPFxPnPListService]';

/** Capture only primitive selector fields; validation belongs to async operations. */
export function snapshotListTarget(target: unknown): ListTargetSnapshot {
  if (typeof target === 'string') {
    return { kind: 'title', value: target };
  }

  if (typeof target !== 'object' || target === null) {
    return { kind: undefined, value: undefined };
  }

  const record = target as Record<string, unknown>;
  const kind = typeof record.kind === 'string' ? record.kind : undefined;
  let value: unknown;

  switch (kind) {
    case 'title': value = record.title; break;
    case 'id': value = record.id; break;
    case 'url': value = record.serverRelativeUrl; break;
    case 'path': value = record.webRelativePath; break;
  }

  return { kind, value: typeof value === 'string' ? value : undefined };
}

function validateRoot(value: string, kind: 'url' | 'path'): void {
  const hasLeadingSlash = value.startsWith('/');
  const wrongBoundary = kind === 'url'
    ? !hasLeadingSlash || value.startsWith('//')
    : hasLeadingSlash;

  if (value.trim().length === 0 || wrongBoundary ||
      /^[a-z][a-z\d+.-]*:/i.test(value.trimStart()) ||
      value.includes('\\') || value.includes('?') ||
      value.split('/').some(segment => segment === '.' || segment === '..')) {
    throw new Error(`${errorPrefix} Invalid ${kind === 'url' ? 'serverRelativeUrl' : 'webRelativePath'}: ` +
      'provide a decoded list root with the correct leading slash, without a URL scheme, query, backslash or dot segments.');
  }
}

function getWebRelativeRoot(sp: SPFI, path: string): string {
  const configuredWebUrl = sp.web.toUrl();
  const endpoint = /\/_api\/web\/?$/i.exec(configuredWebUrl);
  const baseError = (): Error => new Error(`${errorPrefix} Path selection requires an explicit HTTP(S) web base. ` +
    'Configure the supplied SPFI with the intended web URL; unbased and indirect rootWeb clients cannot resolve a path without HTTP.');

  if (!/^https?:\/\//i.test(configuredWebUrl) || !endpoint) {
    throw baseError();
  }

  try {
    // Normalize only the API endpoint casing for PnP's case-sensitive extraction.
    const configuredBase = extractWebUrl(configuredWebUrl.slice(0, endpoint.index) + '/_api/web');
    const webUrl = new URL(configuredBase);
    if ((webUrl.protocol !== 'https:' && webUrl.protocol !== 'http:') || webUrl.search || webUrl.hash) {
      throw baseError();
    }
    const webPath = decodeURIComponent(webUrl.pathname).replace(/\/+$/, '');
    return `${webPath}/${path}`;
  } catch {
    throw baseError();
  }
}

/** Resolve against the actual client that will execute or queue the operation. */
export function resolveListTarget(sp: SPFI, target: ListTargetSnapshot): IList {
  const { kind, value } = target;
  if (value === undefined) {
    throw new Error(`${errorPrefix} Invalid list selector: provide a supported kind and its required string value.`);
  }

  switch (kind) {
    case 'title':
      return sp.web.lists.getByTitle(value);
    case 'id': {
      const trimmed = value.trim();
      const id = trimmed.startsWith('{') && trimmed.endsWith('}') ? trimmed.slice(1, -1) : trimmed;
      if (!/^[a-f\d]{8}-[a-f\d]{4}-[a-f\d]{4}-[a-f\d]{4}-[a-f\d]{12}$/i.test(id)) {
        throw new Error(`${errorPrefix} Invalid list ID: provide a hyphenated GUID, optionally enclosed in paired braces.`);
      }
      return sp.web.lists.getById(id.toLowerCase());
    }
    case 'url':
      validateRoot(value, kind);
      return sp.web.getList(value);
    case 'path':
      validateRoot(value, kind);
      return sp.web.getList(getWebRelativeRoot(sp, value));
    default:
      throw new Error(`${errorPrefix} Invalid list selector kind: use title, id, url or path.`);
  }
}
