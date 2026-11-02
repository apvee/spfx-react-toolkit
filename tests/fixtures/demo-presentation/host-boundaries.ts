const page = { site: { absoluteUrl: 'https://simulated.invalid/sites/demo' }, web: { absoluteUrl: 'https://simulated.invalid/sites/demo' }, cultureInfo: { currentCultureName: 'en-US', currentUICultureName: 'en-US' } };
const forbiddenWrite = async (): Promise<never> => { throw new Error('Fixture forbids CRUD'); };
const unavailable = { isReady: false, isInitializing: false, isLoading: false, error: undefined, invoke: async () => undefined, baseUrl: page.web.absoluteUrl };
const store = { isReady: false, canWrite: false, isLoading: false, isWriting: false, get: async () => undefined, list: async () => [], save: forbiddenWrite, remove: forbiddenWrite };
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
