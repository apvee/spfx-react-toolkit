/** Stable, sanitized categories shared by the protocol, session and outer Field. */
export type FileFolderPickerErrorCode = 'auth-resource' | 'auth-type' | 'auth-challenge' | 'auth-provider'
  | 'selection-identity' | 'selection-scope' | 'selection-policy' | 'selection-unavailable' | 'selection-invalid'
  | 'configuration' | 'opening' | 'onedrive-unavailable' | 'timeout' | 'communication';
const messages: Record<FileFolderPickerErrorCode, string> = {
  configuration: 'Configurazione del selettore non valida. Verifica sorgente, tipi, categoria e limiti; usa il preset senza opzioni legacy.',
  'auth-resource': 'Il selettore ha richiesto una risorsa di autenticazione non consentita dalla configurazione.',
  'auth-type': 'Il selettore ha richiesto un tipo di autenticazione non supportato dalla configurazione.',
  'auth-challenge': 'Il selettore richiede una verifica di autenticazione aggiuntiva non supportata.',
  'auth-provider': 'Non è stato possibile ottenere il token SharePoint. Riprova; se il problema persiste, verifica l’accesso e i permessi esistenti.',
  'selection-identity': 'Non è stato possibile verificare l’identità dell’elemento selezionato. Riapri il selettore e riprova.',
  'selection-scope': 'L’elemento selezionato non appartiene al sito, alla raccolta o alla cartella consentiti.',
  'selection-policy': 'La selezione non rispetta il tipo, la categoria, le estensioni o il numero di elementi consentiti.',
  'selection-unavailable': 'L’elemento selezionato non è disponibile o non è accessibile.',
  'selection-invalid': 'Non è stato possibile verificare la selezione. Nessun nuovo elemento è stato aggiunto.',
  'onedrive-unavailable': 'OneDrive dell’utente corrente non è disponibile o non è accessibile. Verifica che il drive esista e riprova.',
  opening: 'Impossibile aprire il selettore. Verifica la configurazione e l’accesso alla sorgente, quindi riprova.',
  timeout: 'Il selettore non ha risposto in tempo. Chiudilo e riprova.',
  communication: 'La comunicazione con il selettore è stata interrotta. Riaprilo e riprova.',
};
export class FileFolderPickerError extends Error {
  constructor(readonly code: FileFolderPickerErrorCode) {
    super(messages[code]);
    this.name = 'FileFolderPickerError';
  }
}
/** Do not expose provider/native/HTTP payload strings through Field or onError. */
export function fileFolderPickerFailure(failure: unknown, fallback: FileFolderPickerErrorCode): FileFolderPickerError {
  if (failure instanceof FileFolderPickerError) return failure;
  if (fallback !== 'selection-invalid') return new FileFolderPickerError(fallback);
  const detail = failure instanceof Error ? failure.message : '';
  if (/identity mismatch|drive metadata mismatch|provenance|ambiguous.*metadata/i.test(detail)) return new FileFolderPickerError('selection-identity');
  if (/outside|another tenant|crossed a drive|untrusted.*endpoint/i.test(detail)) return new FileFolderPickerError('selection-scope');
  if (/not allowed|selection exceeds|limit|extension|category/i.test(detail)) return new FileFolderPickerError('selection-policy');
  if (/deleted|unavailable|\b403\b|\b404\b/i.test(detail)) return new FileFolderPickerError('selection-unavailable');
  return new FileFolderPickerError(fallback);
}
