export { DisplayMode, ThemeProvider } from '../web-part-shell-browser/spfx-services';
export class SPPermission {
  public constructor(public readonly value: { High: number; Low: number }) {}
  public hasPermission(permission: SPPermission): boolean { return this.value === SPPermission.fullMask.value || ((this.value.Low & permission.value.Low) === permission.value.Low) && ((this.value.High & permission.value.High) === permission.value.High); }
  public hasAllPermissions(...permissions: SPPermission[]): boolean { return permissions.every(permission => this.hasPermission(permission)); }
  static deleteListItems = new SPPermission({ High: 0, Low: 0x8 });
  static emptyMask = new SPPermission({ High: 0, Low: 0 });
  static fullMask = new SPPermission({ High: 0x7fffffff, Low: 0xffffffff });
  static viewListItems = new SPPermission({ High: 0, Low: 1 });
  static editListItems = new SPPermission({ High: 0, Low: 4 });
  static addListItems = new SPPermission({ High: 0, Low: 2 });
  static manageWeb = new SPPermission({ High: 0, Low: 0x40000000 });
  static manageLists = new SPPermission({ High: 0, Low: 0x800 });
}
export class PageContext { static serviceKey = 'trim-page-context'; }
export class HttpClient { static serviceKey = 'clients-http'; static configurations = { v1: {} }; }
export class SPHttpClient { static configurations = { v1: {} }; static serviceKey = 'trim-sp-http'; }
export class MSGraphClientFactory { static serviceKey = 'trim-graph-factory'; }
export class AadHttpClientFactory { static serviceKey = 'clients-aad-factory'; }
export class AadHttpClient { static configurations = { v1: {} }; }

// Native factory and token work are controlled separately from business requests.
export class AadTokenProviderFactory { static serviceKey = 'fixture-aad-token-factory'; }
