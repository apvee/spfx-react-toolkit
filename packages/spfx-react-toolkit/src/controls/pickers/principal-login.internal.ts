/** Only the documented SharePoint membership claim wrapper is equivalent to its UPN login alias. */
export const principalLoginKey = (login: string): string => login.replace(/^i:0#\.f\|membership\|/i, '').toLowerCase();
