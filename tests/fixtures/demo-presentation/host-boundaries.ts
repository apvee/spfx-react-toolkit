const forbiddenWrite = async (): Promise<never> => { throw new Error('Fixture forbids CRUD'); };
const list = { items: [], loading: false, isEmpty: false, hasMore: false, query: async () => [], loadMore: async () => [], clearError: () => {}, create: forbiddenWrite, update: forbiddenWrite, remove: forbiddenWrite };
export const useSPFxPnPList = () => list;
