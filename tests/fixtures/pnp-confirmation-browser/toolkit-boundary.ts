import * as React from 'react';
interface Item { readonly Id: number; readonly Title: string; }
interface Removal { readonly id: number; readonly listTitle: string; readonly operation: number; }
const calls: Removal[] = [];
let settleRemoval: (() => void) | undefined;
let replaceOperation: (() => void) | undefined;
let publishRows: (() => void) | undefined;
let removalCompletion: Promise<void> | undefined;
const pendingRefresh: number[] = [];
declare global { interface Window { pnpConfirmationMock: { readonly calls: readonly Removal[]; settle(): Promise<void>; publishRows(): void; replaceOperation(): void; }; } }
window.pnpConfirmationMock = { calls, settle: async () => { settleRemoval?.(); await removalCompletion; }, publishRows: () => publishRows?.(), replaceOperation: () => replaceOperation?.() };
/** UI-only transport boundary; never creates a PnP client or sends a request. */
export function useSPFxPnPList<T>(listTitle: string, _options?: { pageSize?: number }, _context?: typeof context): {
  items: T[]; loading: boolean; loadingMore: boolean; error: undefined; isEmpty: boolean; hasMore: boolean;
  query: (build: (query: { select(...fields: string[]): { orderBy(field: string, ascending: boolean): void } }) => void) => Promise<T[]>;
  loadMore: () => Promise<T[]>; clearError: () => void;
  create: (data: { Title: string }) => Promise<number>; update: (id: number, data: { Title: string }) => Promise<void>; remove: (id: number) => Promise<void>;
} {
  const [items, setItems] = React.useState<Item[]>([{ Id: 42, Title: 'Intended item' }, { Id: 73, Title: 'Other item' }]);
  const [operation, setOperation] = React.useState(0);
  React.useEffect(() => {
    replaceOperation = () => setOperation(value => value + 1);
    publishRows = () => {
      const ids = pendingRefresh.splice(0);
      setItems(current => current.filter(item => !ids.includes(item.Id)));
    };
    return () => { replaceOperation = undefined; publishRows = undefined; };
  }, []);
  const forbidden = async (): Promise<never> => { throw new Error('Other writes forbidden in confirmation fixture'); };
  return {
    // The actual demo requests exactly these fields; this mock has no network data boundary.
    items: items as unknown as T[], loading: false, loadingMore: false, error: undefined, isEmpty: items.length === 0, hasMore: false,
    query: async () => items as unknown as T[], loadMore: async () => [], clearError: () => {}, create: forbidden, update: forbidden,
    remove: id => {
      calls.push({ id, listTitle, operation });
      // Match the real hook: mutation resolves before a later debounced query publishes rows.
      removalCompletion = new Promise<void>(resolve => { settleRemoval = resolve; }).then(() => { pendingRefresh.push(id); });
      return removalCompletion;
    },
  };
}

// Whole-hook-doubled UI conformance: stable identity, no provider/network proof.
const context = { sp: {}, isInitialized: true, error: undefined, siteUrl: 'https://confirmation.fixture.invalid/sites/demo' };
export const useSPFxPnPContext = (): typeof context => context;
export { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
export * from '../../../packages/spfx-react-toolkit/src/helpers/styles';
