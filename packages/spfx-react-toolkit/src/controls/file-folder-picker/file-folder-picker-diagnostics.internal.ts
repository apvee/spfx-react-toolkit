/** Internal finite diagnostics. Neither native payloads nor provider exceptions may escape here. */
export type FilePickerAuthenticationStage = 'resource-missing' | 'resource-invalid-url' | 'resource-invalid-shape'
  | 'audience-denied' | 'type-unsupported' | 'type-audience-mismatch' | 'claims-denied'
  | 'accepted' | 'provider-failed' | 'token-supplied';
export interface FilePickerAuthenticationDiagnostic {
  readonly stage: FilePickerAuthenticationStage;
  readonly resource: 'source' | 'graph' | 'tenant-personal-candidate' | 'tenant-content-candidate' | 'other' | 'missing' | 'invalid';
  readonly type: 'SharePoint' | 'SharePoint_SelfIssued' | 'SharePoint_MSAL' | 'Graph' | 'missing' | 'unsupported';
  readonly nonRootPath?: boolean;
  readonly hasQuery?: boolean;
  readonly hasFragment?: boolean;
  readonly hasCredentials?: boolean;
  readonly https?: boolean;
}
export function describeFilePickerAuthentication(stage: FilePickerAuthenticationStage, data: Readonly<Record<string, unknown>>, allowedOrigins: readonly string[]): FilePickerAuthenticationDiagnostic {
  const type = data.type === undefined ? 'missing'
    : data.type === 'SharePoint' || data.type === 'SharePoint_SelfIssued' || data.type === 'SharePoint_MSAL' || data.type === 'Graph' ? data.type : 'unsupported';
  if (typeof data.resource !== 'string') return { stage, resource: 'missing', type };
  let url: URL;
  try { url = new URL(data.resource); } catch { return { stage, resource: 'invalid', type }; }
  let resource: FilePickerAuthenticationDiagnostic['resource'] = allowedOrigins.includes(url.origin) ? 'source' : url.origin === 'https://graph.microsoft.com' ? 'graph' : 'other';
  // A relationship label is diagnostic only; it never admits a token or a metadata endpoint.
  if (resource === 'other') for (const origin of allowedOrigins) {
    const host = new URL(origin).hostname;
    if (host.endsWith('.sharepoint.com') && !host.endsWith('-my.sharepoint.com') && url.hostname === host.replace('.sharepoint.com', '-my.sharepoint.com')) resource = 'tenant-personal-candidate';
    if (host.endsWith('-my.sharepoint.com') && url.hostname === host.replace('-my.sharepoint.com', '.sharepoint.com')) resource = 'tenant-content-candidate';
  }
  return { stage, resource, type, nonRootPath: url.pathname !== '/', hasQuery: !!url.search, hasFragment: !!url.hash, hasCredentials: !!(url.username || url.password), https: url.protocol === 'https:' };
}
/** Existing SPFx debug query is an explicit local opt-in. No logger user/site enrichment or storage. */
export function filePickerDebugObserver(document: Document): ((event: FilePickerAuthenticationDiagnostic) => void) | undefined {
  if (new URLSearchParams(document.location.search).get('debug') !== 'true') return undefined;
  return event => { console.debug('[FileFolderPicker authentication]', JSON.stringify(event)); };
}
