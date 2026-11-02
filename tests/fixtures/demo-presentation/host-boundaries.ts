const forbiddenWrite = async (): Promise<never> => { throw new Error('Fixture forbids CRUD'); };
const list = { items: [], loading: false, isEmpty: false, hasMore: false, query: async () => [], loadMore: async () => [], clearError: () => {}, create: forbiddenWrite, update: forbiddenWrite, remove: forbiddenWrite };
export const useSPFxPnPList = () => list;
const suggest = async (query: string) => [`${query} simulated suggestion`];
const search = { results: [], refiners: [], totalResults: 0, loading: false, hasMore: false, search: async () => [], searchWithBuilder: async () => [], applyRefiner: async () => [], clearError: () => {}, suggest };
export const useSPFxPnPSearch = () => search;
export const SearchVerticals = { People: 'simulated-people', Documents: 'simulated-documents', Pages: 'simulated-pages', Videos: 'simulated-videos' };
