import { resolveEditableTextLocaleName } from '../../editable-text/localization/editable-text-locales.internal';
import type { WebPartShellStrings } from './web-part-shell-strings.internal';
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
export const webPartShellLocaleCatalogs: Readonly<Record<string, WebPartShellStrings>> = Object.freeze({
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

/** Use the shared culture policy for title names; never read browser or regional state. */
export function resolveWebPartShellStrings(uiCulture: string | undefined): WebPartShellStrings {
  return webPartShellLocaleCatalogs[resolveEditableTextLocaleName(uiCulture)];
}
