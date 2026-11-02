// Simulation only: preserve real context/list hooks, services and PnP request
// construction. Replace only HTTP send; this cannot establish tenant access.
import { createSPFxPnPContextService as createRealService } from '../../../packages/spfx-react-toolkit/lib/services/spfx-pnp-context.service';

type Request = { readonly url: string; readonly method: string };
const calls: Request[] = [];
let failure: string | undefined;
let hold = false;
const pending: Array<() => void> = [];

declare global {
  interface Window {
    pnpReadFixture: {
      readonly simulation: true;
      readonly calls: readonly Request[];
      configure(next: { failure?: string; hold?: boolean }): void;
      release(): void;
    };
  }
}
window.pnpReadFixture = {
  simulation: true,
  get calls() { return calls; },
  configure(next) { failure = next.failure; hold = next.hold ?? hold; },
  release() { hold = false; pending.splice(0).forEach(complete => complete()); },
};

export const createSPFxPnPContextService: typeof createRealService = (...args) => {
  const service = createRealService(...args);
  return {
    ...service,
    createSPFI: (...options) => service.createSPFI(...options).using(instance => {
      instance.on.send.replace(async (url, init) => {
        // New context cases use the shared native terminal queue, never pilot list JSON.
        // The existing builder resolves this helper's factory to the selected source/lib mode.
        const target = new URL(String(url));
        const read = /\/_api\/web(?:\/lists|\/currentUser)?$/.test(target.pathname);
        const selectorRead = /\/_api\/web\/(?:lists\('[a-f\d-]+'\)|getList\('.+'\))\/items$/.test(target.pathname);
        const protocol = /\/_api\/(?:contextinfo|\$batch)$/.test(target.pathname);
        if (read || selectorRead || protocol) {
          if ((read || selectorRead) && (init.method ?? 'GET') !== 'GET') throw new Error('Simulated read fixture forbids CRUD');
          // Only the bounded terminal protocol may accept digest/read-batch POST.
          // These routes never alter the original GET-only pilot ledger or policy.
          return fetch(String(url), init);
        }
        const method = init.method ?? 'GET';
        calls.push({ url: String(url), method });
        if (method !== 'GET') throw new Error('Simulated read fixture forbids CRUD');
        const error = failure;
        if (hold) await new Promise<void>(resolve => pending.push(resolve));
        else await new Promise<void>(resolve => setTimeout(resolve, 60));
        if (error) throw new Error(`Simulated request error: ${error}`);
        return new Response(JSON.stringify({ value: [{ Id: 1, Title: 'Simulated list row (no tenant request)' }] }), {
          headers: { 'Content-Type': 'application/json' },
        });
      });
      return instance;
    }),
  };
};
