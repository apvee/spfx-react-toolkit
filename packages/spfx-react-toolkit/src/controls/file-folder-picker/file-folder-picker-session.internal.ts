import { resolveFileFolderPickerStrings } from './file-folder-picker-strings.internal';
import type { FileFolderPickerStringsOverride } from './file-folder-picker-strings.types';
import { FileFolderPickerError, fileFolderPickerFailure } from './file-folder-picker-errors.internal';
import type { FileFolderPickerOption } from './file-folder-picker.types';
import { isFilePickerInitialize, readFilePickerCommand, readFilePickerTokenRequest, isFilePickerPageLoaded } from './file-folder-picker-protocol.internal';
import type { FilePickerTokenRequest } from './file-folder-picker-protocol.internal';
import { describeFilePickerAuthentication } from './file-folder-picker-diagnostics.internal';
import type { FilePickerAuthenticationDiagnostic } from './file-folder-picker-diagnostics.internal';

export type FileFolderPickerSessionResult =
  | { readonly status: 'picked'; readonly items: readonly FileFolderPickerOption[] }
  | { readonly status: 'cancelled' };
export interface FileFolderPickerSessionOptions {
  readonly strings?: FileFolderPickerStringsOverride;
  readonly ownerWindow: Window;
  readonly contentWindow: Window;
  readonly trustedOrigin: string;
  readonly channelId: string;
  readonly allowedResourceOrigins: readonly string[];
  readonly graphEnabled?: boolean;
  readonly claimsChallengeEnabled?: boolean;
  readonly onReady?: () => void;
  readonly requestToken: (request: FilePickerTokenRequest) => Promise<string>;
  readonly validatePick: (items: readonly unknown[]) => Promise<readonly FileFolderPickerOption[]>;
  /** Iframes are owned by React/Stage; their Window must not be closed imperatively. */
  readonly closeWindow?: boolean;
  /** Current actionable validation errors; authentication/provider details remain sanitized. */
  readonly onError?: (error: Error) => void;
  readonly onAuthenticationSuccess?: () => void;
  readonly onAuthenticationDiagnostic?: (event: FilePickerAuthenticationDiagnostic) => void;
  /** Trusted current-user discovery only; never reused as selectable metadata origins. */
  readonly resolveAuxiliaryAuthenticationOrigins?: () => Promise<readonly string[]>;
  readonly handshakeTimeoutMs?: number;
  readonly commandTimeoutMs?: number;
  readonly closePollMs?: number;
}
export interface FileFolderPickerSession {
  run(): Promise<FileFolderPickerSessionResult>;
  cancel(): void;
  dispose(): void;
}

/** One opening owns one channel/port and at most one accepted confirmation. No token or payload logging. */
export function createFileFolderPickerSession(options: FileFolderPickerSessionOptions): FileFolderPickerSession {
  const strings = resolveFileFolderPickerStrings(options.strings);
  let active = true, started = false, confirming = false, ready = false, portEpoch = 0;
  const challengedResources = new Set<string>();
  const diagnosticKeys = new Set<string>();
  let auxiliaryOrigins: Promise<readonly string[]> | undefined;
  function diagnostic(event: FilePickerAuthenticationDiagnostic, epoch: number): void {
    if (!active || epoch !== portEpoch || options.onAuthenticationDiagnostic === undefined || diagnosticKeys.size >= 16) return;
    const key = JSON.stringify(event);
    if (diagnosticKeys.has(key)) return;
    diagnosticKeys.add(key);
    try { options.onAuthenticationDiagnostic(event); } catch { /* Observer failures never affect token delivery. */ }
  }
  let port: MessagePort | undefined;
  let handshakeTimer: ReturnType<typeof setTimeout> | undefined;
  let closeTimer: ReturnType<typeof setInterval> | undefined;
  const commandTimers = new Set<ReturnType<typeof setTimeout>>();
  const commandIds = new Set<string | number>();
  let resolveResult: (result: FileFolderPickerSessionResult) => void = () => undefined;
  let rejectResult: (error: Error) => void = () => undefined;
  const result = new Promise<FileFolderPickerSessionResult>((resolve, reject) => { resolveResult = resolve; rejectResult = reject; });
  function cleanup(): void {
    active = false;
    options.ownerWindow.removeEventListener('message', initialize);
    if (handshakeTimer !== undefined) clearTimeout(handshakeTimer);
    if (closeTimer !== undefined) clearInterval(closeTimer);
    commandTimers.forEach(clearTimeout); commandTimers.clear();
    if (port !== undefined) {
      // MessagePort's platform contract requires null to clear its event handler.
      port.onmessage = null;
      port.close(); port = undefined;
    }
    if (options.closeWindow !== false && !options.contentWindow.closed) options.contentWindow.close();
  }
  function releaseConfirmation(): void { confirming = false; }
  function cancel(): void { if (active) { cleanup(); resolveResult({ status: 'cancelled' }); } }
  function fail(error: Error): void { if (active) { cleanup(); rejectResult(error); } }
  function post(message: unknown): void { if (active && port !== undefined) port.postMessage(message); }
  function initialize(event: MessageEvent<unknown>): void {
    if (!active || event.source !== options.contentWindow || event.origin !== options.trustedOrigin ||
      !isFilePickerInitialize(event.data, options.channelId) || event.ports.length !== 1) return;
    if (port !== undefined) {
      port.onmessage = null; // Legacy platform event-handler absence.
      port.close();
      commandTimers.forEach(clearTimeout); commandTimers.clear(); commandIds.clear(); confirming = false;
    }
    const epoch = ++portEpoch;
    if (handshakeTimer !== undefined) clearTimeout(handshakeTimer);
    port = event.ports[0];
    port.onmessage = event => { command(event.data, epoch).catch(() => fail(new FileFolderPickerError('communication', strings))); };
    port.start(); post({ type: 'activate' });
  }
  async function command(input: unknown, epoch: number): Promise<void> {
    if (!active || epoch !== portEpoch) return;
    if (isFilePickerPageLoaded(input)) {
      if (!ready) { ready = true; options.onReady?.(); }
      return;
    }
    const envelope = readFilePickerCommand(input);
    if (envelope === undefined || commandIds.has(envelope.id)) return;
    commandIds.add(envelope.id);
    post({ type: 'acknowledge', id: envelope.id });
    if (envelope.data.command === 'close') { post({ type: 'result', id: envelope.id, data: { result: 'success' } }); cancel(); return; }
    // Classification controls reporting/timeout ownership only; it never authorizes a request.
    const authentication = envelope.data.command === 'authenticate';
    const audience = authentication ? describeFilePickerAuthentication('accepted', envelope.data, options.allowedResourceOrigins).resource : undefined;
    const optionalAuthentication = authentication && (audience === 'graph' || audience === 'tenant-personal-candidate');
    let expired = false;
    const current = (): boolean => active && epoch === portEpoch && !expired;
    const reply = (data: unknown): void => { if (current()) post(data); };
    const timer = setTimeout(() => {
      if (!current()) return;
      if (!optionalAuthentication) { fail(new FileFolderPickerError('timeout', strings)); return; }
      reply({ type: 'result', id: envelope.id, data: { result: 'error', error: {
        code: 'unableToObtainToken', message: strings.protocol.optionalAuthenticationTimeout,
      } } });
      expired = true;
      commandTimers.delete(timer);
    }, options.commandTimeoutMs ?? 120000);
    commandTimers.add(timer);
    let success = false, ownsConfirmation = false;
    let tokenRequest: FilePickerTokenRequest | undefined;
    let auxiliaryAuthentication = false;
    try {
      switch (envelope.data.command) {
        case 'authenticate': {
          let request: FilePickerTokenRequest;
          let rejection: FilePickerAuthenticationDiagnostic | undefined;
          try {
            request = readFilePickerTokenRequest(envelope.data, options.allowedResourceOrigins, { ...options, onDiagnostic: event => { rejection = event; diagnostic(event, epoch); } });
          } catch (failure) {
            const personalRequest = rejection?.stage === 'audience-denied' && rejection.resource === 'tenant-personal-candidate'
              && (rejection.type === 'SharePoint' || rejection.type === 'SharePoint_MSAL' || rejection.type === 'SharePoint_SelfIssued');
            if (!personalRequest || options.resolveAuxiliaryAuthenticationOrigins === undefined) throw failure;
            auxiliaryAuthentication = true;
            // Validate the remaining command schema before discovery. This provisional origin never
            // reaches the provider; only the independently verified result below can authorize it.
            readFilePickerTokenRequest(envelope.data, [...options.allowedResourceOrigins, new URL(String(envelope.data.resource)).origin], { ...options, onDiagnostic: event => { if (event.stage !== 'accepted') diagnostic(describeFilePickerAuthentication(event.stage, envelope.data, options.allowedResourceOrigins), epoch); } });
            const resolveOrigins = options.resolveAuxiliaryAuthenticationOrigins;
            // Discovery belongs to the opening; a replacement port can await the same queued work.
            auxiliaryOrigins ??= Promise.resolve().then(() => {
              if (!active) throw new FileFolderPickerError('auth-resource', strings);
              return resolveOrigins();
            });
            let origins: readonly string[];
            try { origins = await auxiliaryOrigins; } catch { throw new FileFolderPickerError('auth-resource', strings); }
            if (!current()) return;
            request = readFilePickerTokenRequest(envelope.data, [...options.allowedResourceOrigins, ...origins], { ...options, onDiagnostic: event => { if (current()) diagnostic(describeFilePickerAuthentication(event.stage, envelope.data, options.allowedResourceOrigins), epoch); } });
          }
          tokenRequest = request;
          if (!current()) return;
          if (request.claims !== undefined) {
            if (challengedResources.has(request.resource)) {
              diagnostic(describeFilePickerAuthentication('claims-denied', envelope.data, options.allowedResourceOrigins), epoch);
              throw new FileFolderPickerError('auth-challenge', strings);
            }
            challengedResources.add(request.resource);
          }
          let token: string;
          try { token = await options.requestToken(request); } catch {
            if (!current()) return;
            diagnostic(describeFilePickerAuthentication('provider-failed', envelope.data, options.allowedResourceOrigins), epoch);
            throw new FileFolderPickerError('auth-provider', strings);
          }
          if (!current()) return;
          if (typeof token !== 'string' || token.length === 0) {
            diagnostic(describeFilePickerAuthentication('provider-failed', envelope.data, options.allowedResourceOrigins), epoch);
            throw new FileFolderPickerError('auth-provider', strings);
          }
          reply({ type: 'result', id: envelope.id, data: { result: 'token', token } });
          diagnostic(describeFilePickerAuthentication('token-supplied', envelope.data, options.allowedResourceOrigins), epoch);
          if (current() && options.allowedResourceOrigins.includes(request.resource)) options.onAuthenticationSuccess?.();
          break;
        }
        case 'pick': {
          if (confirming) throw new Error('Another selection is being verified.');
          if (!Array.isArray(envelope.data.items) || envelope.data.items.length === 0) throw new Error('The picker returned no items.');
          confirming = true; ownsConfirmation = true;
          const items = await options.validatePick(envelope.data.items);
          if (active && epoch === portEpoch) {
            reply({ type: 'result', id: envelope.id, data: { result: 'success' } });
            success = true; cleanup(); resolveResult({ status: 'picked', items });
          }
          break;
        }
        default:
          reply({ type: 'result', id: envelope.id, data: { result: 'error', error: { code: 'unsupportedCommand', message: strings.protocol.unsupportedCommand } } });
      }
    } catch (failure) {
      // Do not reflect native payloads/provider error strings: they can contain claims or other sensitive data.
      if (!current()) return;
      const error = fileFolderPickerFailure(failure, authentication ? 'auth-provider' : 'selection-invalid', strings);
      // Microsoft may request optional audiences (for example MySite) even with their pivots disabled.
      // Deny that individual request without claiming authentication of our configured source failed.
      const optionalFailure = optionalAuthentication || auxiliaryAuthentication || (authentication && tokenRequest !== undefined && !options.allowedResourceOrigins.includes(tokenRequest.resource));
      if (active && epoch === portEpoch && error.code !== 'auth-resource' && !optionalFailure) options.onError?.(error);
      reply({ type: 'result', id: envelope.id, data: { result: 'error', error: {
        code: authentication ? 'unableToObtainToken' : 'unusableItem',
        message: audience === 'graph' ? strings.protocol.optionalGraphAuthenticationFailed : optionalFailure ? strings.protocol.optionalPersonalAuthenticationFailed : error.message,
      } } });
    } finally {
      if (ownsConfirmation && !success && epoch === portEpoch) releaseConfirmation();
      clearTimeout(timer); commandTimers.delete(timer);
    }
  }
  return {
    run() {
      if (started || !active) return result;
      started = true;
      options.ownerWindow.addEventListener('message', initialize);
      handshakeTimer = setTimeout(() => fail(new FileFolderPickerError('timeout', strings)), options.handshakeTimeoutMs ?? 30000);
      closeTimer = setInterval(() => { if (options.contentWindow.closed) cancel(); }, options.closePollMs ?? 500);
      return result;
    },
    cancel, dispose: cancel,
  };
}
