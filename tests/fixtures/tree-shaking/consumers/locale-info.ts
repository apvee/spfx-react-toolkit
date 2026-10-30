import { getSPFxLocaleInfo } from '__TOOLKIT__';
export function mappedZone(context: Parameters<typeof getSPFxLocaleInfo>[0]): string | undefined { return getSPFxLocaleInfo(context).ianaTimeZone; }
