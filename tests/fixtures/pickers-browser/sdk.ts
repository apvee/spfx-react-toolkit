// Only host SDK service identifiers/enums are simulated; provider/hooks/controls are compiled production modules.
export const keys = { page: {}, sp: {}, graph: {}, theme: {} };
export const SPHttpClient = { serviceKey: keys.sp, configurations: { v1: {} } };
export const MSGraphClientFactory = { serviceKey: keys.graph };
export const PageContext = { serviceKey: keys.page };
export const ThemeProvider = { serviceKey: keys.theme };
export const DisplayMode = { Read: 1, Edit: 2 };
export const PropertyPaneFieldType = { Custom: 1 };
