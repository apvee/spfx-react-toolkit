export { DisplayMode, ThemeProvider } from '../web-part-shell-browser/spfx-services';
export class SPPermission {
  public constructor(public readonly value: { High: number; Low: number }) {}
  public hasPermission(permission: SPPermission): boolean { return ((this.value.Low & permission.value.Low) === permission.value.Low) && ((this.value.High & permission.value.High) === permission.value.High); }
  static viewListItems = new SPPermission({ High: 0, Low: 1 });
  static editListItems = new SPPermission({ High: 0, Low: 4 });
  static addListItems = new SPPermission({ High: 0, Low: 2 });
  static manageWeb = new SPPermission({ High: 0, Low: 0x40000000 });
  static manageLists = new SPPermission({ High: 0, Low: 0x800 });
}
export class PageContext { static serviceKey = 'trim-page-context'; }
export class HttpClient { static configurations = { v1: {} }; }
export class SPHttpClient { static configurations = { v1: {} }; static serviceKey = 'trim-sp-http'; }
export class MSGraphClientFactory { static serviceKey = 'trim-graph-factory'; }
export class AadHttpClient { static configurations = { v1: {} }; }

// Native authentication transport is intentionally unavailable in this local fixture.
export class AadTokenProviderFactory { static serviceKey = 'fixture-aad-token-factory'; }
