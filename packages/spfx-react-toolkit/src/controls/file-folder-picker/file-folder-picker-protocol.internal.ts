import { FileFolderPickerError } from './file-folder-picker-errors.internal';
import { describeFilePickerAuthentication } from './file-folder-picker-diagnostics.internal';
import type { FilePickerAuthenticationDiagnostic, FilePickerAuthenticationStage } from './file-folder-picker-diagnostics.internal';
/** The owned subset of Microsoft's v8 message protocol; incoming structured data remains unknown. */
export interface FilePickerCommandEnvelope {
  readonly type: 'command';
  readonly id: string | number;
  readonly data: Readonly<Record<string, unknown>> & { readonly command: string };
}
function plain(input: unknown): input is Record<string, unknown> {
  // Native postMessage interop may receive null; absence in toolkit state remains undefined.
  if (typeof input !== 'object' || input === null || Array.isArray(input) || Object.getPrototypeOf(input) !== Object.prototype) return false;
  return Object.values(Object.getOwnPropertyDescriptors(input)).every(descriptor => 'value' in descriptor);
}
export function isFilePickerInitialize(input: unknown, channelId: string): boolean {
  return plain(input) && input.type === 'initialize' && input.channelId === channelId;
}
export function readFilePickerCommand(input: unknown): FilePickerCommandEnvelope | undefined {
  if (!plain(input) || input.type !== 'command' ||
    !((typeof input.id === 'string' && input.id.length > 0 && input.id.length <= 256) || (typeof input.id === 'number' && Number.isSafeInteger(input.id))) ||
    !plain(input.data) || typeof input.data.command !== 'string') return undefined;
  // Complete discriminator and members checked above, preserving unknown command-specific fields for their guards.
  return input as unknown as FilePickerCommandEnvelope;
}
export interface FilePickerTokenRequest {
  readonly resource: string;
  readonly type?: 'SharePoint' | 'SharePoint_SelfIssued' | 'SharePoint_MSAL' | 'Graph';
  /** Validated ephemeral claims; never logged or stored beyond this command. */
  readonly claims?: string;
}
/** Exact source audiences remain authoritative; Graph and claims are independent opt-ins. */
export function readFilePickerTokenRequest(data: Readonly<Record<string, unknown>>, allowedOrigins: readonly string[], options: { readonly graphEnabled?: boolean; readonly claimsChallengeEnabled?: boolean; readonly onDiagnostic?: (event: FilePickerAuthenticationDiagnostic) => void } = {}): FilePickerTokenRequest {
  const diagnostic = (stage: FilePickerAuthenticationStage): void => {
    try { options.onDiagnostic?.(describeFilePickerAuthentication(stage, data, allowedOrigins)); } catch { /* Diagnostics never affect authorization. */ }
  };
  const reject = (stage: FilePickerAuthenticationStage, code: 'auth-resource' | 'auth-type' | 'auth-challenge'): never => { diagnostic(stage); throw new FileFolderPickerError(code); };
  if (typeof data.resource !== 'string') return reject('resource-missing', 'auth-resource');
  let url: URL;
  try { url = new URL(data.resource); } catch { return reject('resource-invalid-url', 'auth-resource'); }
  const graph = url.origin === 'https://graph.microsoft.com';
  if (url.protocol !== 'https:' || url.username || url.password || url.search || url.hash || url.pathname !== '/') return reject('resource-invalid-shape', 'auth-resource');
  if (!(graph ? options.graphEnabled === true : allowedOrigins.includes(url.origin))) return reject('audience-denied', 'auth-resource');
  const type = data.type;
  if (type !== undefined && type !== 'Graph' && type !== 'SharePoint' && type !== 'SharePoint_SelfIssued' && type !== 'SharePoint_MSAL') return reject('type-unsupported', 'auth-type');
  if ((graph && type !== 'Graph') || (!graph && type === 'Graph')) return reject('type-audience-mismatch', 'auth-type');
  if (data.claims !== undefined) return reject('claims-denied', 'auth-challenge');
  let claims: string | undefined;
  if (data.claimsChallenge !== undefined) {
    if (!options.claimsChallengeEnabled || !plain(data.claimsChallenge)) return reject('claims-denied', 'auth-challenge');
    const value = data.claimsChallenge.claims;
    if (typeof value !== 'string' || value.length === 0 || value.length > 65536) return reject('claims-denied', 'auth-challenge');
    try { if (!plain(JSON.parse(value))) throw new Error('Invalid claims.'); } catch { return reject('claims-denied', 'auth-challenge'); }
    claims = value;
  }
  diagnostic('accepted');
  return { resource: url.origin, type, ...(claims === undefined ? {} : { claims }) };
}
/** Informational notifications never receive command acknowledgements or expose unvalidated payloads. */
export function isFilePickerPageLoaded(input: unknown): boolean {
  return plain(input) && input.type === 'notification' && plain(input.data) && input.data.notification === 'page-loaded';
}
