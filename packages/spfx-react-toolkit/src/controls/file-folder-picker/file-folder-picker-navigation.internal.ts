import type { FilePickerV8Configuration } from './file-folder-picker-config.internal';

/** Hidden POST is the hosted protocol's navigation contract, not an alternate HTTP/token service. */
export function postFileFolderPicker(document: Document, targetName: string, baseUrl: string, config: FilePickerV8Configuration, token: string, locale: string): void {
  const url = new URL(`${baseUrl.replace(/\/$/, '')}/_layouts/15/FilePicker.aspx`);
  url.searchParams.set('filePicker', JSON.stringify(config));
  url.searchParams.set('locale', locale);
  const form = document.createElement('form');
  form.method = 'POST'; form.action = url.href; form.target = targetName; form.hidden = true;
  const input = document.createElement('input');
  input.type = 'hidden'; input.name = 'access_token'; input.value = token;
  form.appendChild(input);
  document.body.appendChild(form);
  try { form.submit(); }
  finally { input.value = ''; form.remove(); }
}
