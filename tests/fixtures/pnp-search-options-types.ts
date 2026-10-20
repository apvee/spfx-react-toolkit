import type { UseSPFxPnPSearchOptions } from '../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch';
import type { SPFxPnPSearchOptions, SPFxPnPSearchRequestOptions } from '../../packages/spfx-react-toolkit/src/services/spfx-pnp-search.service';

const hook: UseSPFxPnPSearchOptions = { refinerOperator: 'or' };
const service: SPFxPnPSearchOptions = { refinerOperator: 'and' };
const request: SPFxPnPSearchRequestOptions = {
  refinerOperator: 'or', refinementFilters: new Map([['Size', ['range(min,100)']]])
};
const historical: UseSPFxPnPSearchOptions = { refiners: 'FileType' };
const noOptions: SPFxPnPSearchRequestOptions = {};
// @ts-expect-error Only native same-facet AND/OR choices are accepted.
const invalidHook: UseSPFxPnPSearchOptions = { refinerOperator: 'xor' };
// @ts-expect-error The service has the same bounded operator contract.
const invalidService: SPFxPnPSearchOptions = { refinerOperator: 'any' };
// @ts-expect-error Request overrides must use a bounded operator too.
const invalidRequest: SPFxPnPSearchRequestOptions = { refinerOperator: true };
void [hook, service, request, historical, noOptions, invalidHook, invalidService, invalidRequest];
