---
name: SPFx Project Instructions
description: Contextual guidance for GitHub Copilot when working with SharePoint Framework (SPFx) v1.21.1 projects
applyTo: '**'
---

# SPFx v1.21.1 Development Guidelines

> **Context**: These instructions guide GitHub Copilot when generating, refactoring, or reviewing code in SharePoint Framework (SPFx) v1.21.1 projects. They define naming conventions, lifecycle patterns, API usage, and critical rules to ensure consistent, secure, and maintainable SPFx solutions.

## Environment Requirements
| Requirement | Value |
|-------------|-------|
| Node.js | **22.x ONLY** |
| TypeScript | 5.3.3+ (ES2022) |
| React | 17.x |
| Debug URL | `https://localhost:4321/temp/build/manifests.js` |

> 📚 [Setup Development Environment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)

## Naming Conventions
```typescript
interface IMyProps { }              // I + PascalCase
class MyWebPart { }                 // PascalCase
const API_URL = '';                 // UPPER_SNAKE_CASE
private _context: Context;          // _prefix for private
type Status = 'pending' | 'done';   // Union types preferred
```

## Component Base Classes
| Type | Base Class |
|------|-----------|
| Web Part | `BaseClientSideWebPart<T>` |
| App Customizer | `BaseApplicationCustomizer<T>` |
| Field Customizer | `BaseFieldCustomizer<T>` |
| Command Set | `BaseListViewCommandSet<T>` |
| Form Customizer | `BaseFormCustomizer<T>` |
| ACE (Viva) | `BaseAdaptiveCardExtension<T, S>` |

> 📚 [Web Parts](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts) | [Extensions](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/overview-extensions) | [ACE](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/viva/overview-viva-connections)

## Web Part Lifecycle (CRITICAL)
```typescript
export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  protected async onInit(): Promise<void> {
    await super.onInit();
    // 1. Initialize services, PnPjs, register dynamic data sources
  }
  public render(): void {
    // 2. Render UI - use ReactDom.render() for React
  }
  protected onPropertyPaneFieldChanged(path: string, oldVal: any, newVal: any): void {
    // 3. Handle property changes
  }
  protected onDispose(): void {
    // 4. ALWAYS cleanup: clearInterval, unsubscribe, ReactDom.unmountComponentAtNode
    super.onDispose();
  }
}
```

> 📚 [Web Part Lifecycle](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/notes-on-solution-building)

## React Pattern (Hooks)
```typescript
export const MyComponent: React.FC<IMyComponentProps> = ({ context, title }) => {
  const [items, setItems] = React.useState<IItem[]>([]);
  const [loading, setLoading] = React.useState(false);
  React.useEffect(() => {
    const load = async () => {
      setLoading(true);
      try { setItems(await fetchItems(context)); } finally { setLoading(false); }
    };
    load();
  }, [context]);
  const filtered = React.useMemo(() => items.filter(i => i.active), [items]);
  if (loading) return <Spinner />;
  return <div>{filtered.map(i => <div key={i.id}>{i.title}</div>)}</div>;
};
```

> 📚 [Use React Hooks](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/use-react-hooks)

## API Clients
| Client | Use Case | Setup |
|--------|----------|-------|
| `SPHttpClient` | SharePoint REST | `context.spHttpClient.get/post()` |
| `MSGraphClientV3` | Microsoft Graph | `await context.msGraphClientFactory.getClient('3')` |
| `AadHttpClient` | Custom Azure AD APIs | `await context.aadHttpClientFactory.getClient(uri)` |
| **PnPjs** | SharePoint + Graph | `spfi().using(SPFx(this.context))` |

### PnPjs Quick Reference
```typescript
this._sp = spfi().using(SPFx(this.context));
this._graph = graphfi().using(GraphSPFx(this.context));
// CRUD
const items = await this._sp.web.lists.getByTitle("List").items.select("Id","Title").filter("Status eq 'Active'").top(10)();
await this._sp.web.lists.getByTitle("List").items.add({ Title: "New" });
await this._sp.web.lists.getByTitle("List").items.getById(1).update({ Title: "Updated" });
await this._sp.web.lists.getByTitle("List").items.getById(1).delete();
```

> 📚 [SPHttpClient](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/connect-to-sharepoint-using-rest-api) | [MSGraphClientV3](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph) | [PnPjs](https://pnp.github.io/pnpjs/)

## Service Scope (DI)
```typescript
export class MyService {
  public static readonly serviceKey = ServiceKey.create<IMyService>('App:MyService', MyService);
  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._sp = serviceScope.consume(SPHttpClient.serviceKey);
      this._page = serviceScope.consume(PageContext.serviceKey);
    });
  }
}
// Consume: this._myService = this.context.serviceScope.consume(MyService.serviceKey);
```

> 📚 [Service Scopes](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/service-scope)

## Dynamic Data (Web Part Communication)
```typescript
// SOURCE: implements IDynamicDataCallables
this.context.dynamicDataSourceManager.initializeSource(this);
public getPropertyDefinitions(): readonly IDynamicDataPropertyDefinition[] {
  return [{ id: 'selectedItem', title: 'Selected Item' }];
}
public getPropertyValue(propertyId: string): any { return this._selectedItem; }
this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedItem');
// CONSUMER: const item = this.properties.itemSource?.tryGetValue();
```

> 📚 [Dynamic Data](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/dynamic-data)

## Theme Support
```typescript
// manifest.json: "supportsThemeVariants": true
this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
this._theme = this._themeProvider.tryGetTheme();
this._themeProvider.themeChangedEvent.add(this, (args) => { this._theme = args.theme; this.render(); });
// Use: this._theme?.semanticColors?.bodyBackground, this._theme?.palette?.themePrimary
```

> 📚 [Theme Support](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/supporting-section-backgrounds)

## Configuration Files
### package-solution.json
```json
{ "solution": { "skipFeatureDeployment": true, "includeClientSideAssets": true,
    "webApiPermissionRequests": [{ "resource": "Microsoft Graph", "scope": "User.Read" }] }}
```
### manifest.json
```json
{ "id": "GUID-never-change", "alias": "MyWebPart",
  "supportedHosts": ["SharePointWebPart", "TeamsTab", "TeamsPersonalApp"], "supportsThemeVariants": true }
```

## Gulp Commands
```bash
gulp serve                           # Dev server
gulp build --ship && gulp bundle --ship && gulp package-solution --ship  # Production
```

## Teams Detection
```typescript
const isInTeams = !!this.context.sdks?.microsoftTeams;
if (isInTeams) { const ctx = await this.context.sdks.microsoftTeams.context; }
```

## Localization
```typescript
import * as strings from 'MyWebPartStrings'; // from src/webparts/[name]/loc/
this.domElement.innerHTML = strings.TitleLabel;
```

## CRITICAL RULES FOR COPILOT

### ✅ ALWAYS
- Import from `@microsoft/sp-*` packages
- Use `I` prefix for interfaces: `IMyProps`, `IMyState`
- Define explicit return types: `public render(): void`
- Handle errors with try-catch + user-friendly messages
- Escape user input: `import { escape } from '@microsoft/sp-lodash-subset'`
- Cleanup in `onDispose()`: intervals, subscriptions, `ReactDom.unmountComponentAtNode`
- Use `MSGraphClientV3` (not deprecated `GraphHttpClient`)
- Add JSDoc comments for public methods

### ❌ NEVER
- Use `any` type → use `unknown` or specific types
- Direct DOM manipulation → use React
- Forget `await` on async calls in `onInit()`
- Create memory leaks (setInterval without cleanup)
- Change manifest `id` after deployment
- Store secrets in code → use Azure Key Vault

### Context Inference
| If Found | Then Use |
|----------|----------|
| `@pnp/sp` in package.json | PnPjs patterns |
| `react` in package.json | React functional components + hooks |
| `teams/` folder | Check `context.sdks.microsoftTeams` |
| `webApiPermissionRequests` | MSGraphClientV3 available |
| `IDynamicDataCallables` | Web part is data source |
| `DynamicProperty<T>` in props | Web part consumes dynamic data |
| `supportsThemeVariants: true` | Use ThemeProvider |
| `componentType: "Library"` | Library component patterns |

### Recommended Patterns
| Scenario | Pattern |
|----------|---------|
| API calls | Service classes with ServiceScope |
| State | React hooks: useState, useEffect, useMemo, useCallback |
| Data sharing | Dynamic Data for cross-web-part |
| Performance | React.lazy() + useMemo() + externals |
| Parallel | Promise.all() |

## Quick Reference
```typescript
// Property Pane
PropertyPaneTextField('title', { label: strings.TitleLabel })
PropertyPaneDropdown('list', { label: strings.ListLabel, options: this._options })

// Fetch items
const resp = await this.context.spHttpClient.get(
  `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${escape(listTitle)}')/items`,
  SPHttpClient.configurations.v1);
if (!resp.ok) throw new Error(strings.ErrorMessage);
return (await resp.json()).value;
```

📚 [SPFx Docs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) | [PnPjs](https://pnp.github.io/pnpjs/) | [Graph API](https://learn.microsoft.com/en-us/graph/api/overview)
