'use strict';

const pagination = 'ApveeSpfxReactToolkitPaginationStrings';
const stage = 'ApveeSpfxReactToolkitStageStrings';
const editor = 'ApveeSpfxReactToolkitEditableTextStrings';

// These are actual webpack entrypoints compiled by the locked SPFx Gulp rig.
// Exported values keep the selected controls live without requiring a tenant.
const scenarios = [
  { name: 'helper-only', source: "export { createScopedSPFxStorageKey as default } from '@apvee/spfx-react-toolkit';", resources: [] },
  ...['root', 'domain', 'control'].map(kind => ({
    name: `pagination-${kind}`,
    source: `export { Pagination as default } from '@apvee/spfx-react-toolkit${kind === 'root' ? '' : kind === 'domain' ? '/controls' : '/controls/pagination'}';`,
    resources: [pagination],
  })),
  { name: 'two-families', source: "import { Pagination } from '@apvee/spfx-react-toolkit/controls/pagination';\nimport { Stage } from '@apvee/spfx-react-toolkit/controls/stage';\nexport default { Pagination, Stage };", resources: [pagination, stage] },
  { name: 'lazy-editor', source: "export default function loadEditor() { return import(/* webpackChunkName: 'lazy-editor' */ '@apvee/spfx-react-toolkit/controls/editable-text'); }", resources: [editor], lazy: true },
  { name: 'custom-mapping', source: "export { Pagination as default } from '@apvee/spfx-react-toolkit/controls/pagination';", resources: [pagination], custom: true },
  ...['it-it', 'ar-sa'].map(locale => ({ name: `pagination-${locale}`, source: "export { Pagination as default } from '@apvee/spfx-react-toolkit/controls/pagination';", resources: [pagination], locale })),
];

const mutations = [
  { name: 'missing-asset', kind: 'asset', expected: /Failed to find default locale|default locale.*(?:missing|not found)|No localized files found/i },
  { name: 'missing-registration', kind: 'registration', expected: /Can't resolve ['"]ApveeSpfxReactToolkitPaginationStrings/ },
  { name: 'unknown-native-id', kind: 'unknown', expected: /Can't resolve ['"]ApveeSpfxReactToolkitUnknownStrings/ },
  { name: 'missing-selected-locale', kind: 'locale', locale: 'it-it', expected: /(?:selected locale \(it-it\) is not present in the localized data for resource "ApveeSpfxReactToolkitPaginationStrings"|Selected locale it-it did not emit ApveeSpfxReactToolkitPaginationStrings)/ },
  { name: 'root-resource-registry', kind: 'registry', expected: /Unexpected manifest resources/ },
].map(mutation => ({ ...mutation, resources: [pagination] }));

module.exports = { scenarios, mutations };
