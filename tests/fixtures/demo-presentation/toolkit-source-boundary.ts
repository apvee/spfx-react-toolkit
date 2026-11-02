// Retain all real feature controls/styles. Only unavailable host and network hooks
// are doubled in host-boundaries; no root compiled barrel pulls private SPFx resources.
export * from '../../../packages/spfx-react-toolkit/src/helpers/styles';
export { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
export { useStableCallback } from '../../../packages/spfx-react-toolkit/src/hooks/useStableCallback';
export { useSPFxFluent9ThemeInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxFluent9ThemeInfo';
export { useSPFxProperties } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxProperties';
export { useSPFxDisplayMode } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxDisplayMode';
export * from '../../../packages/spfx-react-toolkit/src/controls/content-card';
export * from '../../../packages/spfx-react-toolkit/src/controls/preview-card';
export * from '../../../packages/spfx-react-toolkit/src/controls/description-list';
export * from '../../../packages/spfx-react-toolkit/src/controls/pagination';
export * from '../../../packages/spfx-react-toolkit/src/controls/data-grid';
export * from '../../../packages/spfx-react-toolkit/src/controls/form';
export * from '../../../packages/spfx-react-toolkit/src/controls/stage';
export * from '../../../packages/spfx-react-toolkit/src/controls/settings-group';
export * from '../../../packages/spfx-react-toolkit/src/controls/accordion-group';
export * from '../../../packages/spfx-react-toolkit/src/controls/state-message';
export * from '../../../packages/spfx-react-toolkit/src/controls/auto-skeleton';
export * from '../../../packages/spfx-react-toolkit/src/controls/relative-time';
export * from '../../../packages/spfx-react-toolkit/src/controls/formatted-text';
export * from '../../../packages/spfx-react-toolkit/src/controls/web-part-shell';
export { SPFxWebPartProvider } from '../../../packages/spfx-react-toolkit/src/core/provider-webpart';
export { SPFxApplicationCustomizerProvider } from '../../../packages/spfx-react-toolkit/src/core/provider-application-customizer';
export { SPFxFieldCustomizerProvider } from '../../../packages/spfx-react-toolkit/src/core/provider-field-customizer';
export { SPFxListViewCommandSetProvider } from '../../../packages/spfx-react-toolkit/src/core/provider-listview-commandset';
export * from './host-boundaries';

export * from '../../../packages/spfx-react-toolkit/src/controls/sharepoint-permission-trim';
export * from '../../../packages/spfx-react-toolkit/src/controls/sharepoint-group-trim';
export * from '../../../packages/spfx-react-toolkit/src/controls/entra-group-trim';
export { useSPFxCrossSitePermissions } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions';
export { useSPFxPermissions } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPermissions';

export * from '../../../packages/spfx-react-toolkit/src/controls/editable-text';
export * from '../../../packages/spfx-react-toolkit/src/controls/pickers';
export * from '../../../packages/spfx-react-toolkit/src/controls/file-folder-picker';
export { useSPFxConfigurationStore } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxConfigurationStore';

// Explicit exports override the unavailable legacy hook stubs for the read pilot.
export { useSPFxPnPContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext';
export { useSPFxPnPList } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPList';

// Runtime cases use the actual provider-bound subject hooks and native context witness.
export { useSPFxContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContext';
export { useSPFxSessionStorage, useSPFxLocalStorage } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage';
export { useSPFxLogger } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxLogger';
export { useSPFxPerformance } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPerformance';

// Context scenarios use actual owning hooks; only native SDK services/transports are doubled.
export { useSPFxPageContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPageContext';
export { useSPFxInstanceInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxInstanceInfo';
export { useSPFxIsEdit } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxDisplayMode';
export { useSPFxEnvironmentInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxEnvironmentInfo';
export { useSPFxUserInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxUserInfo';
export { useSPFxSiteInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxSiteInfo';
export { useSPFxListInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxListInfo';
export { useSPFxLocaleInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxLocaleInfo';
export { useSPFxHubSiteInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxHubSiteInfo';
export { useSPFxTeams } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxTeams';
export { useSPFxPageType } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPageType';
export { useSPFxCorrelationInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxCorrelationInfo';
export { useSPFxContainerInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContainerInfo';
export { useSPFxContainerSize } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContainerSize';
export { useSPFxThemeInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxThemeInfo';
export { useSPFxServiceScope } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxServiceScope';
