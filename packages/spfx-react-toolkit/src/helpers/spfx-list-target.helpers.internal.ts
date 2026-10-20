/** Normalize a list GUID while retaining the caller's established error prefix. */
export function normalizeListId(value: string, errorPrefix: string): string {
  const trimmed = value.trim();
  const id = trimmed.startsWith('{') && trimmed.endsWith('}') ? trimmed.slice(1, -1) : trimmed;
  if (!/^[a-f\d]{8}-[a-f\d]{4}-[a-f\d]{4}-[a-f\d]{4}-[a-f\d]{12}$/i.test(id)) {
    throw new Error(`${errorPrefix} Invalid list ID: provide a hyphenated GUID, optionally enclosed in paired braces.`);
  }
  return id.toLowerCase();
}
/** Validate a decoded list root; neutral and free of PnP registrations. */
export function validateListRoot(value: string, kind: 'url' | 'path', errorPrefix: string): void {
  const wrongBoundary = kind === 'url' ? !value.startsWith('/') || value.startsWith('//') : value.startsWith('/');
  if (value.trim().length === 0 || wrongBoundary || /^[a-z][a-z\d+.-]*:/i.test(value.trimStart()) ||
      value.includes('\\') || value.includes('?') || value.split('/').some(segment => segment === '.' || segment === '..')) {
    throw new Error(`${errorPrefix} Invalid ${kind === 'url' ? 'serverRelativeUrl' : 'webRelativePath'}: ` +
      'provide a decoded list root with the correct leading slash, without a URL scheme, query, backslash or dot segments.');
  }
}
/** Combine an explicit web URL with a previously validated decoded root. */
export function webRelativeListRoot(webUrl: URL, path: string): string {
  return `${decodeURIComponent(webUrl.pathname).replace(/\/+$/, '')}/${path}`;
}
