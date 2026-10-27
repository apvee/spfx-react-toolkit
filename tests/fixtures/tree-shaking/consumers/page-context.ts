import { getSPFxSiteInfo } from '__TOOLKIT__';
export function site(context: Parameters<typeof getSPFxSiteInfo>[0]): string { return getSPFxSiteInfo(context).webUrl; }
