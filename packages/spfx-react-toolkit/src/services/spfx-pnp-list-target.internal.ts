import { normalizeListId, validateListRoot, webRelativeListRoot } from '../helpers/spfx-list-target.helpers.internal';
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
    return webRelativeListRoot(webUrl, path);
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
      return sp.web.lists.getById(normalizeListId(value, errorPrefix));
    }
    case 'url':
      validateListRoot(value, kind, errorPrefix);
      return sp.web.getList(value);
    case 'path':
      validateListRoot(value, kind, errorPrefix);
      return sp.web.getList(getWebRelativeRoot(sp, value));
    default:
      throw new Error(`${errorPrefix} Invalid list selector kind: use title, id, url or path.`);
  }
}
