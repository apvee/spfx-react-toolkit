import * as React from 'react';
const page = { site: { absoluteUrl: 'https://simulated.invalid/sites/demo' }, web: { absoluteUrl: 'https://simulated.invalid/sites/demo' }, cultureInfo: { currentCultureName: 'en-US', currentUICultureName: 'en-US' } };
const forbiddenWrite = async (): Promise<never> => { throw new Error('Fixture forbids CRUD'); };
const unavailable = { isReady: false, isInitializing: false, isLoading: false, error: undefined, invoke: async () => undefined, baseUrl: page.web.absoluteUrl };
function useLocalRead() {
  const [isLoading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<Error>();
  const load = React.useCallback(async () => { setLoading(true); await new Promise(resolve => setTimeout(resolve, 100)); setError(new Error('Simulated network unavailable')); setLoading(false); }, []);
  return { isReady: false, isLoading, error, load, reload: load, isWriting: false, write: forbiddenWrite };
}
export const useSPFxUserPhoto = useLocalRead;
export const useSPFxOneDriveAppData = useLocalRead;
export const useSPFxTenantProperty = useLocalRead;
const store = { isReady: false, canWrite: false, isLoading: false, isWriting: false, get: async () => undefined, list: async () => [], save: forbiddenWrite, remove: forbiddenWrite };
export const useSPFxSiteKeyValueStore = () => store;
export const useSPFxTenantKeyValueStore = () => store;
export const useSPFxPnPContext = () => ({ isInitialized: false, sp: undefined });
// Match the real hook's stable callback identity across renders of a scope.
const pnp = { ...unavailable, batch: async () => undefined, clearError: () => {} };
export const useSPFxPnP = () => pnp;
const list = { items: [], loading: false, isEmpty: false, hasMore: false, query: async () => [], loadMore: async () => [], clearError: () => {}, create: forbiddenWrite, update: forbiddenWrite, remove: forbiddenWrite };
export const useSPFxPnPList = () => list;
export const useSPFxPnPListById = () => list;
export const useSPFxPnPListByUrl = () => list;
export const useSPFxPnPListByPath = () => list;
const suggest = async (query: string) => [`${query} simulated suggestion`];
const search = { results: [], refiners: [], totalResults: 0, loading: false, hasMore: false, search: async () => [], searchWithBuilder: async () => [], applyRefiner: async () => [], clearError: () => {}, suggest };
export const useSPFxPnPSearch = () => search;
export const SearchVerticals = { People: 'simulated-people', Documents: 'simulated-documents', Pages: 'simulated-pages', Videos: 'simulated-videos' };
