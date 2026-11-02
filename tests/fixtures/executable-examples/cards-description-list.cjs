const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const assert = require('node:assert/strict');
const directory = path.resolve(__dirname, '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples');
/** Read the actual emitted catalogue without invoking any lazy scenario loader. */
function scenario(title, ownerKey) {
  const module = { exports: {} };
  const compiled = ts.transpileModule(fs.readFileSync(path.join(directory, 'catalogue.ts'), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 },
  });
  new Function('module', 'exports', compiled.outputText)(module, module.exports);
  const entries = module.exports.executableExamples.filter(item => item.title === title && (ownerKey === undefined || item.ownerKey === ownerKey));
  assert.equal(entries.length, 1, `Exactly one actual source-owned example: ${ownerKey ?? ''}/${title}`);
  const entry = entries[0];
  const file = path.join(directory, entry.id + '.tsx');
  return { entry, file, source: fs.readFileSync(file, 'utf8') };
}
module.exports = { scenario };
