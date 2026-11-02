// Actual compiled modules: do not replace ThemeInfo/Teams hooks in this branch.
export * from '../../../packages/spfx-react-toolkit/lib/controls/web-part-shell';
export { useSPFxProperties } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxProperties';
export { useSPFxContext } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxContext';
export { useSPFxDisplayMode } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxDisplayMode';
export { useSPFxFluent9ThemeInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxFluent9ThemeInfo';
export { useSx } from '../../../packages/spfx-react-toolkit/lib/hooks/useSx';
export * from '../../../packages/spfx-react-toolkit/lib/helpers/styles';

export * from '../../../packages/spfx-react-toolkit/lib/controls/editable-text';

export { useSPFxPnPContext } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxPnPContext';
export { useSPFxPnPList } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxPnPList';

// The real catalogue also emits the import-interoperability scenario.
export { useStableCallback } from '../../../packages/spfx-react-toolkit/lib/hooks/useStableCallback';
export { Pagination, type PaginationChangeData, type PaginationProps, type PaginationSize, type PaginationState } from '../../../packages/spfx-react-toolkit/lib/controls/pagination';

// The actual catalogue is typechecked and bundled even when this host selects Shell.
export * from '../../../packages/spfx-react-toolkit/lib/controls/stage';
export * from '../../../packages/spfx-react-toolkit/lib/controls/form';
export * from '../../../packages/spfx-react-toolkit/lib/controls/content-card';
export * from '../../../packages/spfx-react-toolkit/lib/controls/preview-card';
export * from '../../../packages/spfx-react-toolkit/lib/controls/description-list';
export * from '../../../packages/spfx-react-toolkit/lib/controls/settings-group';
export * from '../../../packages/spfx-react-toolkit/lib/controls/accordion-group';
export * from '../../../packages/spfx-react-toolkit/lib/controls/state-message';
export * from '../../../packages/spfx-react-toolkit/lib/controls/auto-skeleton';
export * from '../../../packages/spfx-react-toolkit/lib/controls/relative-time';
export * from '../../../packages/spfx-react-toolkit/lib/controls/formatted-text';

// All generated cases are bundled by the Shell conformance fixture.
export { useSPFxSessionStorage, useSPFxLocalStorage } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxStorage';
export { useSPFxLogger } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxLogger';
export { useSPFxPerformance } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxPerformance';

// Prior generated DataGrid cases are also dependencies of this complete catalogue.
export * from '../../../packages/spfx-react-toolkit/lib/controls/data-grid';

// The complete canonical catalogue includes the actual Context cases.
export { useSPFxPageContext } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxPageContext';
export { useSPFxInstanceInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxInstanceInfo';
export { useSPFxIsEdit } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxDisplayMode';
export { useSPFxEnvironmentInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxEnvironmentInfo';
export { useSPFxUserInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxUserInfo';
export { useSPFxSiteInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxSiteInfo';
export { useSPFxListInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxListInfo';
export { useSPFxLocaleInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxLocaleInfo';
export { useSPFxHubSiteInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxHubSiteInfo';
export { useSPFxTeams } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxTeams';
export { useSPFxPageType } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxPageType';
export { useSPFxCorrelationInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxCorrelationInfo';
export { useSPFxPermissions } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxPermissions';
export { useSPFxCrossSitePermissions, type SPFxCrossSitePermissionsOptions } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxCrossSitePermissions';
export { useSPFxContainerInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxContainerInfo';
export { useSPFxContainerSize } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxContainerSize';
export { useSPFxThemeInfo } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxThemeInfo';
export { useSPFxServiceScope } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxServiceScope';
export type { SPFxScopedPermissionsOptions } from '../../../packages/spfx-react-toolkit/lib/services/spfx-permission-target.types';

// These actual Clients/GraphData subjects are dependencies of the complete lazy catalogue.
export { useSPFxHttpClient } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxHttpClient';
export { useSPFxSPHttpClient } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxSPHttpClient';
export { useSPFxMSGraphClient } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxMSGraphClient';
export { useSPFxAadHttpClient } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxAadHttpClient';
export { useSPFxAadTokenProvider } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxAadTokenProvider';
export { useSPFxApiPermissionPrecheck } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxApiPermissionPrecheck';
export { useSPFxUserPhoto } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxUserPhoto';
export { useSPFxOneDriveAppData } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxOneDriveAppData';
export { summarizeSPFxApiPermissionResults, type SPFxApiPermissionCheckResult } from '../../../packages/spfx-react-toolkit/lib/helpers/spfx-api-permission-precheck.helpers';

// Site cases call the actual public hook and factory; only native HTTP is controlled.
export { useSPFxSiteKeyValueStore, type SPFxSiteKeyValueStoreItem } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxSiteKeyValueStore';
export { createSPFxSiteKeyValueStoreService, type SPFxSiteKeyValueStoreServiceItem } from '../../../packages/spfx-react-toolkit/lib/services/spfx-site-key-value-store.service';

// Tenant property cases retain actual hooks, discovery and the direct read-only factory.
export { useSPFxTenantProperty, type SPFxTenantPropertyResult } from '../../../packages/spfx-react-toolkit/lib/hooks/useSPFxTenantProperty';
export { createSPFxTenantPropertyService, type SPFxTenantProperty, type SPFxTenantPropertyService } from '../../../packages/spfx-react-toolkit/lib/services/spfx-tenant-property.service';
