import type { EditableTextStrings } from './editable-text-strings.internal';
import { strings as arSAStrings } from './locales/ar-SA';
import { strings as azLatnAZStrings } from './locales/az-Latn-AZ';
import { strings as bgBGStrings } from './locales/bg-BG';
import { strings as bsLatnBAStrings } from './locales/bs-Latn-BA';
import { strings as caESStrings } from './locales/ca-ES';
import { strings as csCZStrings } from './locales/cs-CZ';
import { strings as cyGBStrings } from './locales/cy-GB';
import { strings as daDKStrings } from './locales/da-DK';
import { strings as deDEStrings } from './locales/de-DE';
import { strings as elGRStrings } from './locales/el-GR';
import { strings as enUSStrings } from './locales/en-US';
import { strings as esESStrings } from './locales/es-ES';
import { strings as etEEStrings } from './locales/et-EE';
import { strings as euESStrings } from './locales/eu-ES';
import { strings as fiFIStrings } from './locales/fi-FI';
import { strings as frFRStrings } from './locales/fr-FR';
import { strings as gaIEStrings } from './locales/ga-IE';
import { strings as glESStrings } from './locales/gl-ES';
import { strings as heILStrings } from './locales/he-IL';
import { strings as hiINStrings } from './locales/hi-IN';
import { strings as hrHRStrings } from './locales/hr-HR';
import { strings as huHUStrings } from './locales/hu-HU';
import { strings as idIDStrings } from './locales/id-ID';
import { strings as itITStrings } from './locales/it-IT';
import { strings as jaJPStrings } from './locales/ja-JP';
import { strings as kkKZStrings } from './locales/kk-KZ';
import { strings as koKRStrings } from './locales/ko-KR';
import { strings as ltLTStrings } from './locales/lt-LT';
import { strings as lvLVStrings } from './locales/lv-LV';
import { strings as mkMKStrings } from './locales/mk-MK';
import { strings as msMYStrings } from './locales/ms-MY';
import { strings as nbNOStrings } from './locales/nb-NO';
import { strings as nlNLStrings } from './locales/nl-NL';
import { strings as plPLStrings } from './locales/pl-PL';
import { strings as prsAFStrings } from './locales/prs-AF';
import { strings as ptBRStrings } from './locales/pt-BR';
import { strings as ptPTStrings } from './locales/pt-PT';
import { strings as roROStrings } from './locales/ro-RO';
import { strings as ruRUStrings } from './locales/ru-RU';
import { strings as skSKStrings } from './locales/sk-SK';
import { strings as slSIStrings } from './locales/sl-SI';
import { strings as srCyrlRSStrings } from './locales/sr-Cyrl-RS';
import { strings as srLatnRSStrings } from './locales/sr-Latn-RS';
import { strings as svSEStrings } from './locales/sv-SE';
import { strings as thTHStrings } from './locales/th-TH';
import { strings as trTRStrings } from './locales/tr-TR';
import { strings as ukUAStrings } from './locales/uk-UA';
import { strings as viVNStrings } from './locales/vi-VN';
import { strings as zhCNStrings } from './locales/zh-CN';
import { strings as zhTWStrings } from './locales/zh-TW';

/** Immutable private catalogs; module evaluation creates only local resource objects. */
export const editableTextLocaleCatalogs: Readonly<Record<string, EditableTextStrings>> = Object.freeze({
  'ar-SA': arSAStrings,
  'az-Latn-AZ': azLatnAZStrings,
  'bg-BG': bgBGStrings,
  'bs-Latn-BA': bsLatnBAStrings,
  'ca-ES': caESStrings,
  'cs-CZ': csCZStrings,
  'cy-GB': cyGBStrings,
  'da-DK': daDKStrings,
  'de-DE': deDEStrings,
  'el-GR': elGRStrings,
  'en-US': enUSStrings,
  'es-ES': esESStrings,
  'et-EE': etEEStrings,
  'eu-ES': euESStrings,
  'fi-FI': fiFIStrings,
  'fr-FR': frFRStrings,
  'ga-IE': gaIEStrings,
  'gl-ES': glESStrings,
  'he-IL': heILStrings,
  'hi-IN': hiINStrings,
  'hr-HR': hrHRStrings,
  'hu-HU': huHUStrings,
  'id-ID': idIDStrings,
  'it-IT': itITStrings,
  'ja-JP': jaJPStrings,
  'kk-KZ': kkKZStrings,
  'ko-KR': koKRStrings,
  'lt-LT': ltLTStrings,
  'lv-LV': lvLVStrings,
  'mk-MK': mkMKStrings,
  'ms-MY': msMYStrings,
  'nb-NO': nbNOStrings,
  'nl-NL': nlNLStrings,
  'pl-PL': plPLStrings,
  'prs-AF': prsAFStrings,
  'pt-BR': ptBRStrings,
  'pt-PT': ptPTStrings,
  'ro-RO': roROStrings,
  'ru-RU': ruRUStrings,
  'sk-SK': skSKStrings,
  'sl-SI': slSIStrings,
  'sr-Cyrl-RS': srCyrlRSStrings,
  'sr-Latn-RS': srLatnRSStrings,
  'sv-SE': svSEStrings,
  'th-TH': thTHStrings,
  'tr-TR': trTRStrings,
  'uk-UA': ukUAStrings,
  'vi-VN': viVNStrings,
  'zh-CN': zhCNStrings,
  'zh-TW': zhTWStrings,
});


// No prefix matching: ambiguous languages and unsupported scripts retain the en-US fallback.
const safeAliases: Readonly<Record<string, string>> = Object.freeze({
  'ar': 'ar-SA',
  'bg': 'bg-BG',
  'ca': 'ca-ES',
  'cs': 'cs-CZ',
  'cy': 'cy-GB',
  'da': 'da-DK',
  'de': 'de-DE',
  'el': 'el-GR',
  'en': 'en-US',
  'es': 'es-ES',
  'et': 'et-EE',
  'eu': 'eu-ES',
  'fi': 'fi-FI',
  'fr': 'fr-FR',
  'ga': 'ga-IE',
  'gl': 'gl-ES',
  'he': 'he-IL',
  'hi': 'hi-IN',
  'hr': 'hr-HR',
  'hu': 'hu-HU',
  'id': 'id-ID',
  'it': 'it-IT',
  'ja': 'ja-JP',
  'kk': 'kk-KZ',
  'ko': 'ko-KR',
  'lt': 'lt-LT',
  'lv': 'lv-LV',
  'mk': 'mk-MK',
  'ms': 'ms-MY',
  'nb': 'nb-NO',
  'nl': 'nl-NL',
  'pl': 'pl-PL',
  'prs': 'prs-AF',
  'ro': 'ro-RO',
  'ru': 'ru-RU',
  'sk': 'sk-SK',
  'sl': 'sl-SI',
  'sv': 'sv-SE',
  'th': 'th-TH',
  'tr': 'tr-TR',
  'uk': 'uk-UA',
  'vi': 'vi-VN',
  'en-gb': 'en-US',
  'en-au': 'en-US',
  'en-ca': 'en-US',
  'en-nz': 'en-US',
  'en-ie': 'en-US',
  'en-za': 'en-US',
  'de-at': 'de-DE',
  'de-ch': 'de-DE',
  'fr-ca': 'fr-FR',
  'fr-be': 'fr-FR',
  'fr-ch': 'fr-FR',
  'nl-be': 'nl-NL',
  'es-mx': 'es-ES',
  'es-ar': 'es-ES',
  'es-cl': 'es-ES',
  'es-co': 'es-ES',
  'es-pe': 'es-ES',
  'zh-hans': 'zh-CN',
  'zh-hans-cn': 'zh-CN',
  'zh-hans-sg': 'zh-CN',
  'zh-sg': 'zh-CN',
  'zh-hant': 'zh-TW',
  'zh-hant-tw': 'zh-TW',
  'zh-hant-hk': 'zh-TW',
  'zh-hant-mo': 'zh-TW',
  'zh-hk': 'zh-TW',
  'zh-mo': 'zh-TW',
  'sr-latn': 'sr-Latn-RS',
  'sr-latn-ba': 'sr-Latn-RS',
  'sr-latn-me': 'sr-Latn-RS',
  'sr-cyrl': 'sr-Cyrl-RS',
  'sr-cyrl-ba': 'sr-Cyrl-RS',
  'sr-cyrl-me': 'sr-Cyrl-RS',
  'az-latn': 'az-Latn-AZ',
  'bs-latn': 'bs-Latn-BA',
});

/** Resolve an explicit culture or safe alias, preserving the existing 50-culture fallback contract. */
export function resolveEditableTextLocaleName(locale?: string): string {
  const normalized = typeof locale === 'string' ? locale.trim().replace(/_/g, '-').toLowerCase() : '';
  const exact = Object.keys(editableTextLocaleCatalogs).find(culture => culture.toLowerCase() === normalized);
  if (exact) return exact;
  return Object.prototype.hasOwnProperty.call(safeAliases, normalized) ? safeAliases[normalized] : 'en-US';
}
/** Resolve neutral control messages without reading browser or SPFx state. */
export function resolveEditableTextStrings(locale?: string): EditableTextStrings {
  return editableTextLocaleCatalogs[resolveEditableTextLocaleName(locale)];
}
