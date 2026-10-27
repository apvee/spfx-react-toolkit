import type { FileFolderPickerSession, FileFolderPickerSessionOptions, FileFolderPickerSessionResult } from '../../../packages/spfx-react-toolkit/lib/controls/file-folder-picker/file-folder-picker-session.internal';
import { rawItem, observations } from './transport';
interface Opening { options: FileFolderPickerSessionOptions; pick(indexes: readonly number[]): Promise<void>; cancel(): void; fail(): void; }
const openings = new Map<string, Opening>();
/** Simulates only the hosted session boundary. validatePick is the production metadata resolver. */
export function createFileFolderPickerSession(options: FileFolderPickerSessionOptions): FileFolderPickerSession {
  observations.sessions++;
  let finish: (result: FileFolderPickerSessionResult) => void = () => undefined;
  let reject: (error: Error) => void = () => undefined;
  let active = true;
  const result = new Promise<FileFolderPickerSessionResult>((resolve, failure) => { finish = resolve; reject = failure; });
  const cancel = (): void => { if (active) { active = false; finish({ status: 'cancelled' }); } };
  const opening: Opening = { options, cancel, fail() { if (active) { active = false; reject(new Error('Simulated hosted failure')); } },
    async pick(indexes) {
      try { const items = await options.validatePick(indexes.map(rawItem)); if (active) { active = false; finish({ status: 'picked', items }); } }
      catch (error) { if (active) { active = false; reject(error instanceof Error ? error : new Error('Simulated confirmation failed')); } }
    },
  };
  openings.set(options.channelId, opening);
  return { run: () => result, cancel, dispose() { observations.disposed++; cancel(); openings.delete(options.channelId); } };
}
/** Replaces the POST/navigation only, so no form/token can reach a remote site. */
export function postFileFolderPicker(_document: Document, name: string): void {
  const opening = openings.get(name);
  if (opening === undefined) throw new Error('Fixture session missing');
  const document = opening.options.contentWindow.document;
  document.open();
  document.write('<!doctype html><html lang="en"><head><meta charset="UTF-8"><style>body{font:16px system-ui;padding:20px;color:#172333;background:#f4f8ff}button{font:inherit;padding:12px;margin:6px;border:1px solid #385b9d;background:white;border-radius:5px}p{max-width:650px}</style></head><body><h1>Simulated Microsoft hosted surface</h1><p>Production FileFolderPicker and Stage surround this iframe. Only hosted navigation/session, SPFx SDK and metadata responses are simulated. No Microsoft service is contacted.</p></body></html>');
  document.close();
  const button = (text: string, action: () => void): void => { const element = document.createElement('button'); element.textContent = text; element.onclick = action; document.body.appendChild(element); };
  button('Confirm file', () => { void opening.pick([0]); });
  button('Confirm PDF', () => { void opening.pick([1]); });
  button('Confirm folder', () => { void opening.pick([2]); });
  button('Confirm all three', () => { void opening.pick([0, 1, 2]); });
  button('Cancel hosted selection', opening.cancel);
  button('Simulate hosted failure', opening.fail);
}
