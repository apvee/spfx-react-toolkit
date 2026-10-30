import { createPickerRequestCache } from '../../controls/pickers/picker-request-cache.internal';
import type { PickerRequestCache } from '../../controls/pickers/picker-request-cache.internal';

interface MountedRuntime { readonly runtime: PickerRequestCache; leases: number; }
export interface PickerRuntimeLease { readonly runtime: PickerRequestCache; release(): void; }
// Local weak lifetime bookkeeping only; no renderer, external registration or persistent cache.
const owners = new WeakMap<object, WeakMap<Document, MountedRuntime>>();
export function acquirePropertyPanePickerRuntime(owner: object, document: Document): PickerRuntimeLease {
  let documents = owners.get(owner);
  if (documents === undefined) { documents = new WeakMap(); owners.set(owner, documents); }
  let mounted = documents.get(document);
  if (mounted === undefined) { mounted = { runtime: createPickerRequestCache(owner), leases: 0 }; documents.set(document, mounted); }
  mounted.leases++;
  const current = mounted, scope = documents;
  let released = false;
  return { runtime: current.runtime, release() {
    if (released) return;
    released = true;
    if (--current.leases === 0) { current.runtime.dispose(); if (scope.get(document) === current) scope.delete(document); }
  } };
}
