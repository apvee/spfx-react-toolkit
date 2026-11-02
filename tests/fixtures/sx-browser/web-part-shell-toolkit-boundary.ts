// Actual source modules: the private SPFx host adapter supplies services, never hooks.
export * from '../../../packages/spfx-react-toolkit/src/controls/web-part-shell';
export { useSPFxProperties } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxProperties';
export { useSPFxContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContext';
export { useSPFxDisplayMode } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxDisplayMode';
export { useSPFxFluent9ThemeInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxFluent9ThemeInfo';
export { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
export * from '../../../packages/spfx-react-toolkit/src/helpers/styles';

export * from '../../../packages/spfx-react-toolkit/src/controls/editable-text';

export { useSPFxPnPContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext';
export { useSPFxPnPList } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPList';

// The real catalogue also emits the import-interoperability scenario.
export { useStableCallback } from '../../../packages/spfx-react-toolkit/src/hooks/useStableCallback';
export { Pagination, type PaginationChangeData, type PaginationProps, type PaginationSize, type PaginationState } from '../../../packages/spfx-react-toolkit/src/controls/pagination';

// The actual catalogue is typechecked and bundled even when this host selects Shell.
export * from '../../../packages/spfx-react-toolkit/src/controls/stage';
export * from '../../../packages/spfx-react-toolkit/src/controls/form';
export * from '../../../packages/spfx-react-toolkit/src/controls/content-card';
export * from '../../../packages/spfx-react-toolkit/src/controls/preview-card';
export * from '../../../packages/spfx-react-toolkit/src/controls/description-list';
export * from '../../../packages/spfx-react-toolkit/src/controls/settings-group';
export * from '../../../packages/spfx-react-toolkit/src/controls/accordion-group';
export * from '../../../packages/spfx-react-toolkit/src/controls/state-message';
export * from '../../../packages/spfx-react-toolkit/src/controls/auto-skeleton';
export * from '../../../packages/spfx-react-toolkit/src/controls/relative-time';
export * from '../../../packages/spfx-react-toolkit/src/controls/formatted-text';

// All generated cases are bundled by the Shell conformance fixture.
export { useSPFxSessionStorage, useSPFxLocalStorage } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage';
export { useSPFxLogger } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxLogger';
export { useSPFxPerformance } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPerformance';

// Prior generated DataGrid cases are also dependencies of this complete catalogue.
export * from '../../../packages/spfx-react-toolkit/src/controls/data-grid';

// The complete canonical catalogue includes the actual Context cases.
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
export { useSPFxPermissions } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPermissions';
export { useSPFxCrossSitePermissions, type SPFxCrossSitePermissionsOptions } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxCrossSitePermissions';
export { useSPFxContainerInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContainerInfo';
export { useSPFxContainerSize } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContainerSize';
export { useSPFxThemeInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxThemeInfo';
export { useSPFxServiceScope } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxServiceScope';
export type { SPFxScopedPermissionsOptions } from '../../../packages/spfx-react-toolkit/src/services/spfx-permission-target.types';

// These actual Clients/GraphData subjects are dependencies of the complete lazy catalogue.
export { useSPFxHttpClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxHttpClient';
export { useSPFxSPHttpClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxSPHttpClient';
export { useSPFxMSGraphClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxMSGraphClient';
export { useSPFxAadHttpClient } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxAadHttpClient';
export { useSPFxAadTokenProvider } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxAadTokenProvider';
export { useSPFxApiPermissionPrecheck } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxApiPermissionPrecheck';
export { useSPFxUserPhoto } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxUserPhoto';
export { useSPFxOneDriveAppData } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxOneDriveAppData';
export { summarizeSPFxApiPermissionResults, type SPFxApiPermissionCheckResult } from '../../../packages/spfx-react-toolkit/src/helpers/spfx-api-permission-precheck.helpers';

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
export { useSPFxConfigurationStore } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxConfigurationStore';
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
