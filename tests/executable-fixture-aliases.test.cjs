const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { executableFixtureAliases } = require('../scripts/executable-fixture-aliases.cjs');
const { executableRoot, ownership } = require('../scripts/executable-examples-render.cjs');
test('browser aliases follow only actual owned imports through exact public source and compiled targets', t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(),'executable-fixture-aliases-'));
  t.after(() => fs.rmSync(root,{recursive:true,force:true}));
  fs.cpSync(path.join(__dirname,'fixtures/executable-examples'),root,{recursive:true});
  fs.symlinkSync(path.join(__dirname,'../node_modules'),path.join(root,'node_modules'),'dir');
  const metadataFile = path.join(root,'package.json');
  const metadata = JSON.parse(fs.readFileSync(metadataFile,'utf8'));
  metadata.exports['./lib/widget'] = { types: './lib/widget.d.ts', default: './lib/widget.js' };
  metadata.exports['./lib/private'] = { types: './lib/private.d.ts', default: './lib/private.js' };
  metadata.exports['./hooks'] = { types: './lib/hooks.d.ts', default: './lib/hooks.js' };
  fs.writeFileSync(path.join(root,'src/hooks.ts'),"export { Widget } from './widget'; export const unrelatedNativeHook = 1;\n");
  fs.writeFileSync(metadataFile,JSON.stringify(metadata));
  fs.writeFileSync(path.join(root,'src/private.ts'),'export const secret = 1;\n');
  fs.mkdirSync(path.join(root,'lib')); fs.writeFileSync(path.join(root,'lib/widget.js'),'exports.Widget = 1;\n');
  const directory = path.join(root,executableRoot); fs.mkdirSync(directory,{recursive:true});
  const module = path.join(directory,'a'.repeat(64)+'.tsx');
  const write = (specifier, marker = ownership) => fs.writeFileSync(module,marker+`\nimport * as entry from '${metadata.name}${specifier}';\n`);
  const config = { packageRoot: '.', sourceRoot: 'src', tsconfig: 'tsconfig.json' };
  write('/lib/widget');
  assert.deepEqual(executableFixtureAliases(root,true,config),{ [metadata.name+'/lib/widget$']:path.join(root,'src/widget.tsx') });
  assert.deepEqual(executableFixtureAliases(root,false,config),{ [metadata.name+'/lib/widget$']:path.join(root,'lib/widget.js') });
  fs.writeFileSync(module,ownership+`\nimport { Widget as Example } from '${metadata.name}/hooks';\n`);
  assert.deepEqual(executableFixtureAliases(root,true,config),{ [metadata.name+'/hooks$']:path.join(root,'src/widget.tsx') }, 'named generic imports use the actual owner without unrelated native hooks');
  assert.deepEqual(executableFixtureAliases(root,false,config),{ [metadata.name+'/hooks$']:path.join(root,'lib/widget.js') });
  fs.writeFileSync(module,ownership+`\nimport * as Hooks from '${metadata.name}/hooks';\n`);
  assert.deepEqual(executableFixtureAliases(root,true,config),{ [metadata.name+'/hooks$']:path.join(root,'src/hooks.ts') }, 'namespace imports retain their complete versioned entry');
  for (const specifier of ['/lib/private','/lib/not-exported','/src/widget','/lib/../widget']) {
    write(specifier); assert.throws(() => executableFixtureAliases(root,true,config),/Unapproved executable fixture import/);
  }
  write('/lib/widget','// Manual file'); assert.throws(() => executableFixtureAliases(root,true,config),/manually owned/);
});
