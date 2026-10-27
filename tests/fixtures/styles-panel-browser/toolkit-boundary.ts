// The sample still imports only the package root. This fixture substitutes only
// the root's unavailable SPFx runtime graph with the same compiled public modules.
export * from '../../../packages/spfx-react-toolkit/lib/helpers/styles';
export { useSx } from '../../../packages/spfx-react-toolkit/lib/hooks/useSx';
export { useStableCallback } from '../../../packages/spfx-react-toolkit/lib/hooks/useStableCallback';
export { useSPFxFluent9ThemeInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxFluent9ThemeInfo';
export { getTeamsFluentTheme } from '../../../packages/spfx-react-toolkit/lib/helpers/spfx-theme.helpers';

export * from '../../../packages/spfx-react-toolkit/lib/controls/formatted-text';
export * from '../../../packages/spfx-react-toolkit/lib/controls/relative-time';
export * from '../../../packages/spfx-react-toolkit/lib/controls/auto-skeleton';
export * from '../../../packages/spfx-react-toolkit/lib/controls/state-message';
export * from '../../../packages/spfx-react-toolkit/lib/controls/settings-group';
export * from '../../../packages/spfx-react-toolkit/lib/controls/accordion-group';
export * from '../../../packages/spfx-react-toolkit/lib/controls/form';
// AutoSkeleton follows the same compiled control boundary used by the real sample.

export { getSPFxLocaleInfo } from '../../../packages/spfx-react-toolkit/lib/helpers/spfx-page-context.helpers';

// The suggestions fixture uses a deterministic service boundary: no search network calls.
const suggest = async (query: string): Promise<string[]> => [`${query} suggestion`];
export const useSPFxPnPSearch = (): { readonly suggest: typeof suggest } => ({ suggest });
