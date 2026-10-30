// Canonical translator data supplied as caller dictionaries in instance-isolation tests.
// Native-loader selection is tested separately against emitted AMD resources.
function editorDictionary(h, culture) {
  const resource = require(`../../../packages/spfx-react-toolkit/src/controls/editable-text/localization/resources/${culture}.json`);
  return h.load('controls/editable-text/localization/editable-text-strings.internal.ts').defineEditableTextStrings(resource);
}
function shellDictionary(culture) {
  return require(`../../../packages/spfx-react-toolkit/src/controls/web-part-shell/localization/resources/${culture}.json`);
}
module.exports = { editorDictionary, shellDictionary };
