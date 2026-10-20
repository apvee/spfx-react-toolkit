import { FileFolderPickerError, fileFolderPickerFailure } from './file-folder-picker-errors.internal';
import * as React from 'react';
import { useFileFolderPickerTheme } from './useFileFolderPickerTheme.internal';
import { useStableCallback } from '../../hooks/useStableCallback';
import { useSPFxAadTokenProvider } from '../../hooks/useSPFxAadTokenProvider';
import { useSPFxPageContext } from '../../hooks/useSPFxPageContext';
import { usePickerSourceContext } from '../pickers/usePickerSourceContext.internal';
import { PickerAuthoritativeEpochContext, PickerRuntimeContext } from '../pickers/picker-runtime-context.internal';
import { createPickerRequestCache } from '../pickers/picker-request-cache.internal';
import { validateFileFolderPickerValues } from './file-folder-picker-values.internal';
import { fileFolderPickerIdentity } from './file-folder-picker-identity.internal';
import { resolveFileFolderPickerScope } from './file-folder-picker-scope.internal';
import { normalizeFileFolderPickerItemPolicy, resolveFileFolderPickerItems, resolveFileFolderPickerValue } from './file-folder-picker-results.internal';
import { compileFileFolderPickerConfig } from './file-folder-picker-config.internal';
import { resolveFileFolderPickerExperience } from './file-folder-picker-presets.internal';
import type { ResolvedFileFolderPickerExperience } from './file-folder-picker-presets.internal';
import { createFileFolderPickerSession } from './file-folder-picker-session.internal';
import type { FileFolderPickerSession } from './file-folder-picker-session.internal';
import { postFileFolderPicker } from './file-folder-picker-navigation.internal';
import { filePickerDebugObserver } from './file-folder-picker-diagnostics.internal';
import type { FileFolderPickerChangeData, FileFolderPickerOption, FileFolderPickerProps, FileFolderPickerValue } from './file-folder-picker.types';

function canonicalConfiguration(input: unknown): string {
  return JSON.stringify(input, (_key, value: unknown) => {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) return value;
    const record = value as Record<string, unknown>;
    return Object.fromEntries(Object.keys(record).sort().map(key => [key, record[key]]));
  });
}
/** Render-time dependency comparison must tolerate invalid JavaScript configuration until Browse validates it. */
function extensionPolicySignature(input: unknown): readonly string[] | 'invalid' | undefined {
  if (input === undefined) return undefined;
  if (!Array.isArray(input)) return 'invalid';
  const extensions: string[] = [];
  for (const extension of input) {
    if (typeof extension !== 'string') return 'invalid';
    extensions.push(extension.toLowerCase());
  }
  return extensions.sort();
}
interface DialogOpening { readonly name: string; readonly epoch: number; readonly experience: ResolvedFileFolderPickerExperience; readonly values: readonly FileFolderPickerValue[]; readonly sourceItems: readonly FileFolderPickerOption[]; readonly theme: 'default' | 'dark'; }
function uniqueValues(input: readonly FileFolderPickerValue[]): readonly FileFolderPickerValue[] {
  const used = new Set<string>();
  return validateFileFolderPickerValues(input).filter(value => { const key = fileFolderPickerIdentity(value); if (used.has(key)) return false; used.add(key); return true; });
}
interface FileFolderPickerRuntime {
  readonly values: readonly FileFolderPickerValue[];
  label(value: FileFolderPickerValue): string;
  extension(value: FileFolderPickerValue): string | undefined;
  remove(value: FileFolderPickerValue): void;
  browse(): void;
  readonly browseButton: React.RefObject<HTMLButtonElement>;
  readonly busy: boolean;
  readonly error?: string;
  readonly hydrating: boolean;
  readonly dialog?: DialogOpening;
  readonly attachFrame: React.RefCallback<HTMLIFrameElement>;
  cancel(): void;
  readonly validation?: string;
  readonly ready: boolean;
  readonly initError?: string;
}
/** Selection ownership and hosted session are separate; hydration never emits or rewrites saved references. */
export function useFileFolderPicker(props: FileFolderPickerProps): FileFolderPickerRuntime {
  const theme = useFileFolderPickerTheme();
  const sourceContext = usePickerSourceContext(), page = useSPFxPageContext(), auth = useSPFxAadTokenProvider();
  const context = React.useMemo(() => ({ ...sourceContext, tokenProvider: auth.tokenProvider }), [sourceContext, auth.tokenProvider]);
  const sharedCache = React.useContext(PickerRuntimeContext), authoritativeEpoch = React.useContext(PickerAuthoritativeEpochContext);
  const localCache = React.useMemo(() => sharedCache ?? createPickerRequestCache(context.owner), [sharedCache, context.owner]);
  React.useEffect(() => () => { if (sharedCache === undefined) localCache.dispose(); }, [localCache, sharedCache]);
  const [controlled] = React.useState(props.value !== undefined);
  const [draft, setDraft] = React.useState<readonly FileFolderPickerValue[]>(() => {
    try { return uniqueValues(props.value ?? props.defaultValue ?? []); } catch { return []; }
  });
  let invalid: string | undefined, values = draft;
  try { if (controlled && props.value !== undefined) values = uniqueValues(props.value); }
  catch { invalid = 'Riferimenti file/cartelle non validi.'; }
  if (controlled !== (props.value !== undefined)) invalid = 'La modalità controllata non può cambiare durante il montaggio.';
  const [resolved, setResolved] = React.useState<readonly FileFolderPickerOption[]>([]);
  const seeds: FileFolderPickerOption[] = [];
  for (const option of props.initialOptions ?? []) {
    try { if (typeof option.text !== 'string' || !option.text.trim()) continue; uniqueValues([option.value]); seeds.push(option); } catch { /* Invalid display seeds never replace saved references. */ }
  }
  const models = new Map([...resolved, ...seeds].map(option => [fileFolderPickerIdentity(option.value), option]));
  const label = (value: FileFolderPickerValue): string => models.get(fileFolderPickerIdentity(value))?.text ?? value.name ?? 'Elemento non risolto';
  const valuesSignature = JSON.stringify(values), policySignature = canonicalConfiguration([props.scope, props.itemTypes ?? 'files', !!props.multiselect, props.selectionLimit, extensionPolicySignature(props.extensions), props.requiredAccess ?? 'read']);
  const session = React.useRef<FileFolderPickerSession>();
  const browseButton = React.useRef<HTMLButtonElement>(null); // React DOM ref contract uses null before mount.
  const restoreFocus = React.useRef(false);
  const epoch = React.useRef(0), mounted = React.useRef(true);
  const [busy, setBusy] = React.useState(false), [error, setError] = React.useState<FileFolderPickerError>(), [hydrating, setHydrating] = React.useState(() => values.length > 0 && values.every(value => !models.has(fileFolderPickerIdentity(value)) && value.name === undefined));
  const [dialog, setDialog] = React.useState<DialogOpening>();
  const cancel = useStableCallback(() => {
    const hadOpening = session.current !== undefined || dialog !== undefined;
    epoch.current++;
    session.current?.dispose(); session.current = undefined;
    if (mounted.current) { if (hadOpening) restoreFocus.current = true; setBusy(false); setDialog(undefined); setError(undefined); }
  });
  React.useLayoutEffect(() => {
    if (!busy && restoreFocus.current) { restoreFocus.current = false; browseButton.current?.focus(); }
  }, [busy]);
  React.useLayoutEffect(() => { cancel(); }, [policySignature, valuesSignature, authoritativeEpoch, context, props.disabled, invalid, cancel]);
  React.useEffect(() => () => { mounted.current = false; cancel(); }, [cancel]);
  const onReady = useStableCallback(() => props.onReady?.());
  const report = useStableCallback((failure: Error) => { if (mounted.current) { const error = fileFolderPickerFailure(failure, 'selection-invalid'); setError(error); props.onError?.(error); } });
  React.useEffect(() => {
    let current = true;
    if (values.length === 0 || invalid !== undefined || context.spClient === undefined) { setHydrating(false); return; }
    setHydrating(values.length > 0 && values.every(value => !models.has(fileFolderPickerIdentity(value)) && value.name === undefined));
    resolveFileFolderPickerScope(props.scope, undefined, context, localCache)
      .then(scope => Promise.all(values.map(value => resolveFileFolderPickerValue(value, scope, context, localCache, props))))
      .then(items => { if (current) setResolved(items.filter((item): item is FileFolderPickerOption => item !== undefined)); })
      .catch(() => { /* Retain saved labels/references; unreadable selections remain removable. */ })
      .finally(() => { if (current) setHydrating(false); });
    return () => { current = false; };
  }, [valuesSignature, policySignature, context, localCache]);
  const emit = useStableCallback((reason: FileFolderPickerChangeData['reason'], next: readonly FileFolderPickerValue[], picked: readonly FileFolderPickerOption[] = []) => {
    if (props.disabled || invalid !== undefined) return;
    const confirmed = uniqueValues(next), map = new Map([...models.values(), ...picked].map(option => [fileFolderPickerIdentity(option.value), option]));
    if (!controlled) setDraft(confirmed);
    setResolved([...map.values()]); setError(undefined);
    props.onChange?.({ reason, values: confirmed, items: confirmed.flatMap(value => { const model = map.get(fileFolderPickerIdentity(value)); return model === undefined ? [] : [model]; }), unresolvedValues: confirmed.filter(value => !map.has(fileFolderPickerIdentity(value))) });
  });
  const remove = useStableCallback((value: FileFolderPickerValue) => { if (!props.disabled && invalid === undefined) emit('remove', values.filter(candidate => fileFolderPickerIdentity(candidate) !== fileFolderPickerIdentity(value))); });
  const run = useStableCallback(async (contentWindow: Window, document: Document, name: string, runEpoch: number, experience: ResolvedFileFolderPickerExperience, openingValues: readonly FileFolderPickerValue[], sourceItems: readonly FileFolderPickerOption[], openingTheme: 'default' | 'dark') => {
    const current = (): boolean => mounted.current && epoch.current === runEpoch;
    const openingProps = experience.options;
    try {
      const scope = await resolveFileFolderPickerScope(openingProps.scope, openingProps.initialLocation, context, localCache, openingProps.initialLocationFallback);
      if (!current()) return;
      const provider = auth.tokenProvider;
      if (provider === undefined) throw new FileFolderPickerError('auth-provider');
      const config = compileFileFolderPickerConfig(scope, { sourceItems }, { origin: document.location.origin, channelId: name, theme: openingTheme }, experience);
      const authenticationDiagnostic = filePickerDebugObserver(document);
      const personalNavigationRequested = scope.kind !== 'oneDrive' && experience.authentication.auxiliaryPersonalResourceEnabled;
      const opened = createFileFolderPickerSession({ ownerWindow: document.defaultView ?? window, contentWindow, trustedOrigin: new URL(scope.baseUrl).origin, channelId: name,
        allowedResourceOrigins: scope.allowedResourceOrigins, graphEnabled: experience.authentication.graphEnabled, claimsChallengeEnabled: experience.authentication.claimsChallengeEnabled, onReady, closeWindow: false, onError: failure => { if (current()) report(failure); },
        onAuthenticationDiagnostic: authenticationDiagnostic === undefined ? undefined : event => { if (current()) authenticationDiagnostic(event); },
        resolveAuxiliaryAuthenticationOrigins: personalNavigationRequested ? async () => {
          if (!current()) throw new FileFolderPickerError('auth-resource');
          // Independently verify the current user's personal drive; native commands cannot introduce trust.
          // Keep these token audiences separate from scope.allowedResourceOrigins and all result checks.
          const personal = await resolveFileFolderPickerScope({ kind: 'oneDrive' }, undefined, context, localCache);
          return [personal.resource];
        } : undefined,
        requestToken: request => request.claims === undefined ? provider.getToken(request.resource) : provider.getToken(request.resource, { claims: request.claims, useCachedToken: false }),
        onAuthenticationSuccess: () => { if (current()) setError(previous => previous?.code.startsWith('auth-') ? undefined : previous); },
        validatePick: items => resolveFileFolderPickerItems(items, scope, context, localCache, { ...openingProps, retainedValues: openingValues, isCurrent: current }),
      });
      session.current = opened;
      const outcome = opened.run();
      // Attach rejection immediately while token acquisition/navigation is pending.
      outcome.catch(() => undefined);
      const beforeNavigation = await Promise.race([
        provider.getToken(scope.resource).then(token => {
          if (!token) throw new FileFolderPickerError('auth-provider');
          return { token };
        }, () => { throw new FileFolderPickerError('auth-provider'); }),
        outcome.then(outcome => ({ outcome })),
      ]);
      if (!current() || 'outcome' in beforeNavigation || contentWindow.closed) return;
      const token = beforeNavigation.token;
      postFileFolderPicker(document, name, scope.baseUrl, config, token, page.cultureInfo.currentUICultureName);
      const result = await outcome;
      if (current() && result.status === 'cancelled') setError(undefined);
      if (current() && result.status === 'picked') emit('pick', openingProps.multiselect ? [...openingValues, ...result.items.map(option => option.value)] : result.items.map(option => option.value), result.items);
    } catch (failure) {
      if (current()) report(fileFolderPickerFailure(failure, 'opening'));
    } finally {
      if (current()) {
        session.current?.dispose(); session.current = undefined;
        restoreFocus.current = true; setBusy(false); setDialog(undefined);
      }
    }
  });
  const browse = useStableCallback(() => {
    if (props.disabled || busy || invalid !== undefined) return;
    let experience: ResolvedFileFolderPickerExperience;
    try { experience = resolveFileFolderPickerExperience(props); }
    catch { report(new FileFolderPickerError('configuration')); return; }
    cancel(); setError(undefined);
    const document = browseButton.current?.ownerDocument ?? window.document, owner = document.defaultView ?? window;
    const crypto = owner.crypto;
    const name = `file-picker-${Array.from(crypto.getRandomValues(new Uint32Array(4))).map(value => value.toString(16)).join('-')}`;
    const runEpoch = epoch.current; setBusy(true);
    // Snapshot configuration before iframe attachment or any asynchronous resolution. Preserve
    // verified model objects: cloning them would discard their private provenance stamps.
    setDialog({ name, epoch: runEpoch, experience, values: [...values], theme, sourceItems: resolved.filter(option => values.some(value => fileFolderPickerIdentity(value) === fileFolderPickerIdentity(option.value))) });
  });
  const attachFrame = React.useCallback((frame: HTMLIFrameElement | null) => {
    // React ref callbacks receive null on detach; no toolkit absence API is widened.
    if (frame?.contentWindow !== null && frame?.contentWindow !== undefined && dialog !== undefined) run(frame.contentWindow, frame.ownerDocument, dialog.name, dialog.epoch, dialog.experience, dialog.values, dialog.sourceItems, dialog.theme).catch(() => undefined);
  }, [dialog, run]);
  const maximum = props.multiselect ? props.selectionLimit : 1;
  let policyFailure: string | undefined;
  try {
    const policy = normalizeFileFolderPickerItemPolicy(props);
    if (values.some(value => (policy.itemTypes === 'files' && value.kind !== 'file') || (policy.itemTypes === 'folders' && value.kind !== 'folder') ||
      (value.kind === 'file' && policy.extensions.length > 0 && value.name !== undefined && !policy.extensions.some(extension => value.name?.toLowerCase().endsWith(extension))))) {
      policyFailure = 'La selezione contiene elementi incompatibili con il tipo o le estensioni configurate. Rimuovili o sostituiscili.';
    }
  } catch { policyFailure = 'Configurazione del tipo, delle estensioni o del limite di selezione non valida.'; }
  const validation = invalid ?? policyFailure ?? (maximum !== undefined && values.length > maximum ? 'La selezione supera il limite configurato. Rimuovi gli elementi in eccesso.' : undefined);
  return { values, label, extension: (value: FileFolderPickerValue): string | undefined => models.get(fileFolderPickerIdentity(value))?.extension ?? value.name?.split('.').pop(), remove, browse, browseButton, busy, error: error?.message, hydrating: hydrating && !values.some(value => models.has(fileFolderPickerIdentity(value)) || value.name !== undefined), dialog, attachFrame, cancel, validation,
    ready: auth.isReady && context.spClient !== undefined, initError: auth.initError === undefined ? undefined : new FileFolderPickerError('auth-provider').message };
}
