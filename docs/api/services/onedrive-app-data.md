# OneDrive App Data Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxOneDriveAppDataService`

```ts
function createSPFxOneDriveAppDataService(
  graphClient: MSGraphClientV3
): SPFxOneDriveAppDataService;
```

Required public types:

- `SPFxOneDriveAppDataReadResult`
- `SPFxOneDriveAppDataService`

Returned methods:

| Method | Description |
|--------|-------------|
| `read<T>(fileName, folder?)` | Reads JSON content from the user's OneDrive app root |
| `write<T>(fileName, content, folder?)` | Writes JSON content to the user's OneDrive app root |

`read` returns `{ data: undefined, isNotFound: true }` for recognized not-found Graph errors. It throws JSON parse errors for invalid stored content and rethrows non-not-found Graph errors.

Reads first obtain complete Graph metadata, without projecting away the download annotation, and immediately fetch the preauthenticated HTTPS `@microsoft.graph.downloadUrl` with `credentials: 'omit'`, without Graph authorization headers. This follows [Microsoft's browser download guidance](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0#downloading-files-in-javascript-apps) and avoids the authenticated `/content` redirect/CORS failure. Missing or invalid download metadata and unsuccessful HTTP downloads are errors. Native `fetch` is required. Writes retain the Graph `/content` endpoint.

## Usage and public contracts

```ts
const data = createSPFxOneDriveAppDataService(graphClient);
const result = await data.read<unknown>('settings.json', 'dashboard');
if (!result.isNotFound) console.log(result.data);
await data.write('settings.json', { theme: 'dark' }, 'dashboard');
```

`SPFxOneDriveAppDataReadResult<T = unknown>` has readonly `data: T | undefined` and `isNotFound: boolean`; T does not validate parsed JSON. `SPFxOneDriveAppDataService.read<T = unknown>(fileName: string, folder?: string)` returns that result; `write<T = unknown>(fileName, content: T, folder?)` returns `Promise<void>` after replacement succeeds. Both use the [Graph path helper](../helpers/graph-path.md): omitted/empty folder selects app root, non-alphanumeric folder characters except hyphen/underscore become hyphens, and filenames are interpolated without encoding or validation. Different folder strings can sanitize to the same namespace.

The caller supplies an initialized Graph client and its app-folder read/write grants. Construction makes no request; no React provider, cache, automatic folder creation or automatic retry is added. Writes call `JSON.stringify` and PUT `application/json` to Graph `/content`; circular/nested bigint serialization and Graph failures reject. JSON-compatible content is required; generic typing does not make undefined/functions/Date/bigint lossless.

Only a recognized Graph metadata 404/itemNotFound code/message produces `isNotFound: true`. Missing or non-HTTPS download URLs, failed/expired downloads (including 404), fetch/CORS failures and invalid JSON all reject. A failed download never becomes absence, preventing a consumer's create-if-missing policy from overwriting an existing file. Download URLs are temporary preauthenticated metadata and must be used immediately without forwarding Graph authorization or cookies.
