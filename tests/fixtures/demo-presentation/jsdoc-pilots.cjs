// Independent browser acceptance expectations, never a generator/navigation input.
const pilotTitles = {
  useSPFxPnP: ['Compare explicit invoke and batch reads'],
  useSPFxPnPContext: ['Read current or cross-site web information'],
  useSPFxConfigurationStore: ['Compare layered configuration and physical inventories'],
  createSPFxConfigurationStoreService: ['Compose tenant defaults and site overrides with the service'],
  useSPFxTenantKeyValueStore: ['Explore tenant records and exact text persistence'],
  createSPFxTenantKeyValueStoreService: ['Round trip numeric looking text through a tenant service'],
  useSPFxTenantProperty: ['Read tenant version and counter with explicit activation'],
  createSPFxTenantPropertyService: ['Read and validate a catalog StorageEntity'],
  useSPFxSiteKeyValueStore: ['Explore collection root records with guarded mutations'],
  createSPFxSiteKeyValueStoreService: ['Read a collection root record before a consented save'],
  useSPFxUserPhoto: ['Read the current user photo with managed URL cleanup'],
  useSPFxOneDriveAppData: ['Read and explicitly replace OneDrive app data'],
  useSPFxHttpClient: ['Load an external todo with tracked invocation state'],
  useSPFxSPHttpClient: ['Read the current SharePoint web title and URL'],
  useSPFxMSGraphClient: ['Read the current Graph profile after initialization'],
  useSPFxAadHttpClient: ['Initialize a custom API client and call a path'],
  useSPFxAadTokenProvider: ['Inspect client readiness and token-provider initialization'],
  useSPFxApiPermissionPrecheck: ['Manually check Graph and custom API delegated scopes'],
  WebPartShell: ['Configure the actual web part Shell', 'Retain content through Shell presentation changes', 'Observe content remount between subtle and Card'],
  Pagination: ['Browse controlled rows', 'Browse controlled rows with pending completion', 'Show an empty controlled result', 'Browse orders across pagination modes'],
  useStableCallback: ['Report current count through a retained callback', 'Retain callback identity through rerenders'],
  useSx: ['Compare root domain and historical imports'],
  useSPFxPnPList: ['Read a user-selected list with observable results'],
  useSPFxContext: ['Inspect the current native host identity'],
  useSPFxPageContext: ['Inspect selected native page metadata'],
  useSPFxInstanceInfo: ['Inspect the provider instance identity'],
  useSPFxDisplayMode: ['Observe the native Read and Edit mode'],
  useSPFxIsEdit: ['Observe the native edit mode shortcut'],
  useSPFxEnvironmentInfo: ['Observe the acquired application host'],
  useSPFxUserInfo: ['Inspect the current user display metadata'],
  useSPFxSiteInfo: ['Inspect the current web and site collection'],
  useSPFxListInfo: ['Inspect the optional current list context'],
  useSPFxLocaleInfo: ['Compare host cultures and seasonal site times'],
  useSPFxHubSiteInfo: ['Start a scope bound hub association lookup'],
  useSPFxTeams: ['Observe application SDK readiness and theme'],
  useSPFxPageType: ['Inspect the supplied page classification'],
  useSPFxCorrelationInfo: ['Inspect current diagnostic identifiers'],
  useSPFxPermissions: ['Inspect native current page permission masks'],
  useSPFxCrossSitePermissions: ['Check staged legacy site web and list permissions', 'Check staged web list item and folder permissions'],
  useSPFxContainerInfo: ['Observe the native container and tracked size'],
  useSPFxContainerSize: ['Observe the measured container width category'],
  useSPFxThemeInfo: ['Observe the native SPFx theme snapshot'],
  useSPFxFluent9ThemeInfo: ['Observe the projected Fluent theme tokens'],
  useSPFxServiceScope: ['Inspect the native registered service scope'],

};
function filterPilotGroups(groups) {
  const features = groups.flatMap(group => group.features);
  for (const [title, scenarios] of Object.entries(pilotTitles)) {
    const matches = features.filter(feature => feature.title === title);
    if (matches.length !== 1) throw new Error(`Missing pilot feature or duplicate: ${title}`);
    for (const scenario of scenarios) if (matches[0].scenarios.filter(entry => entry.title === scenario).length !== 1)
      throw new Error(`Missing pilot scenario or duplicate: ${title}/${scenario}`);
  }
  return groups.map(group => ({ ...group, features: group.features.filter(feature => Object.hasOwn(pilotTitles, feature.title))
    .map(feature => ({ ...feature, scenarios: feature.scenarios.filter(entry => pilotTitles[feature.title].includes(entry.title)) })) }))
    .filter(group => group.features.length > 0);
}
module.exports = { pilotTitles, filterPilotGroups };
