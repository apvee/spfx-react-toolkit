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

// Clients subjects use actual hooks; factories/transports only are native doubles.
export { useSPFxHttpClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxHttpClient';
export { useSPFxSPHttpClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxSPHttpClient';
export { useSPFxMSGraphClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxMSGraphClient';
export { useSPFxAadHttpClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxAadHttpClient';
export { useSPFxAadTokenProvider } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxAadTokenProvider';
export { useSPFxApiPermissionPrecheck } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxApiPermissionPrecheck';
export { summarizeSPFxApiPermissionResults } from '../../../packages/spfx-react-toolkit/src/helpers/spfx-api-permission-precheck.helpers';

// GraphData subjects retain actual hook/service/URL ownership; native transports alone are controlled.
export { useSPFxUserPhoto } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxUserPhoto';
export { useSPFxOneDriveAppData } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxOneDriveAppData';

// Site cases call the actual public hook and factory; only native HTTP is controlled.
export { useSPFxSiteKeyValueStore, type SPFxSiteKeyValueStoreItem } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore';
export { createSPFxSiteKeyValueStoreService, type SPFxSiteKeyValueStoreServiceItem } from '../../../packages/spfx-react-toolkit/src/services/spfx-site-key-value-store.service';

// Tenant property cases retain actual hooks, discovery and the direct read-only factory.
export { useSPFxTenantProperty, type SPFxTenantPropertyResult } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantProperty';
export { createSPFxTenantPropertyService, type SPFxTenantProperty, type SPFxTenantPropertyService } from '../../../packages/spfx-react-toolkit/src/services/spfx-tenant-property.service';

// Tenant list cases retain actual hook/discovery/list implementations.
export { useSPFxTenantKeyValueStore, type SPFxTenantKeyValueStoreItem } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore';
export { createSPFxTenantKeyValueStoreService, type SPFxTenantKeyValueStoreService, type SPFxTenantKeyValueStoreServiceItem } from '../../../packages/spfx-react-toolkit/src/services/spfx-tenant-key-value-store.service';

// Configuration cases retain actual facade, physical stores and mode-specific identity.
export { createSPFxConfigurationStoreService, type SPFxConfigurationStoreItem, type SPFxConfigurationStoreMode, type SPFxConfigurationStoreScope } from '../../../packages/spfx-react-toolkit/src/services/spfx-configuration-store.service';

export { useSPFxPnP } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnP';

// Selector cases execute their actual matching-mode hooks and shared list service.
export { useSPFxPnPListById } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPListById';
export { useSPFxPnPListByUrl } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPListByUrl';
export { useSPFxPnPListByPath } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPListByPath';

// Direct service subjects use actual matching-mode factories, never pilot response doubles.
export { createSPFxPnPContextService } from '../../../packages/spfx-react-toolkit/src/services/spfx-pnp-context.service';
export { createSPFxPnPService } from '../../../packages/spfx-react-toolkit/src/services/spfx-pnp.service';
export { createSPFxPnPListService, type SPFxPnPListSelector, type SPFxPnPListService, type SPFxPnPListQueryBuilder, type SPFxPnPListQueryResult } from '../../../packages/spfx-react-toolkit/src/services/spfx-pnp-list.service';

// Search cases retain the actual matching-mode hook, constants, service and PnP parser.
export { useSPFxPnPSearch, SearchVerticals } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPSearch';

export { getSPFxCorrelationInfo } from '../../../packages/spfx-react-toolkit/src/helpers/spfx-page-context.helpers';
