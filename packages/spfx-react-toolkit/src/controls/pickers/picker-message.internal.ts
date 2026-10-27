import type { PickersStrings } from './pickers-strings.types';
import { resolvePickersStrings } from './pickers-strings.internal';

type PickerOwnedMessage =
  | { readonly group: 'errors'; readonly key: keyof PickersStrings['errors'] }
  | { readonly group: 'status'; readonly key: keyof PickersStrings['status'] };

function messageText(message: PickerOwnedMessage, strings: PickersStrings): string {
  return message.group === 'errors' ? strings.errors[message.key] : strings.status[message.key];
}

/** Stable private ownership marker; raw caller/backend Error objects remain untouched. */
export class PickerMessageError extends Error {
  constructor(readonly descriptor: PickerOwnedMessage) {
    super(messageText(descriptor, resolvePickersStrings()));
    this.name = 'PickerMessageError';
  }
}

/** Project cached toolkit messages with the current instance dictionary only. */
export function projectPickerError(error: Error | undefined, strings: PickersStrings): Error | undefined {
  return error instanceof PickerMessageError ? new Error(messageText(error.descriptor, strings)) : error;
}
