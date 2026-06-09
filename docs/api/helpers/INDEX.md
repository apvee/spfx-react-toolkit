# Helpers API Reference

> Pure helpers for SPFx context mapping, permission checks, storage keys, serialization, Graph paths, and theme conversion.

Helpers are pure functions. They do not read React context, do not perform I/O, and do not manage provider state.

## Helper Groups

| Helper group | Exports |
|--------------|---------|
| Page context mapping | `getSPFxUserInfo`, `getSPFxSiteInfo`, `getSPFxListInfo`, `getSPFxLocaleInfo`, `getSPFxEnvironmentInfo`, `getSPFxPageTypeInfo`, `getSPFxCorrelationInfo` |
| Permissions | `hasSPFxPermission` |
| Container | `getSPFxContainerSize` |
| Storage | `createScopedSPFxStorageKey` |
| Tenant values | `escapeODataValue`, `serializeTenantValue`, `deserializeTenantValue` |
| Graph paths | `buildOneDriveAppDataPath`, `buildUserPhotoEndpoint` |
| Theme | `createFluent9ThemeFromSPFxTheme`, `getTeamsFluentTheme` |

## Usage Pattern

```ts
const user = getSPFxUserInfo(pageContext);
const key = createScopedSPFxStorageKey(instanceId, 'filters');
const endpoint = buildUserPhotoEndpoint({ email, size: '240x240' });
```

Use helpers directly when you already have the required inputs. Use hooks when a React component should read from the SPFx provider.
