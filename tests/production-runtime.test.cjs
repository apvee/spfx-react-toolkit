const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { executeProbe, runRuntimeMatrix } = require('../scripts/verify-production-runtime.cjs');

test('Node-executed production UI uses the browser context through legacy node-conditioned peers', async () => {
  const webpack = require('webpack');
  const { buildRuntimeConfig } = require('../scripts/verify-production-runtime.cjs');
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'runtime-module-context-'));
  try {
    const shared = path.join(directory, 'node_modules/shared-context');
    const owner = path.join(directory, 'node_modules/legacy-provider');
    fs.mkdirSync(shared, { recursive: true }); fs.mkdirSync(owner, { recursive: true });
    fs.writeFileSync(path.join(shared, 'package.json'), JSON.stringify({ name: 'shared-context', exports: { import: './esm.js', require: './commonjs.cjs' } }));
    fs.writeFileSync(path.join(shared, 'esm.js'), 'export const context = {};');
    fs.writeFileSync(path.join(shared, 'commonjs.cjs'), 'exports.context = {};');
    fs.writeFileSync(path.join(owner, 'package.json'), JSON.stringify({ name: 'legacy-provider', main: './commonjs.cjs', module: './esm.js', exports: { '.': { node: './commonjs.cjs', import: './esm.js', require: './commonjs.cjs' } } }));
    fs.writeFileSync(path.join(owner, 'esm.js'), "export { context } from 'shared-context';");
    fs.writeFileSync(path.join(owner, 'commonjs.cjs'), "exports.context = require('shared-context').context;");
    const entry = path.join(directory, 'entry.js');
    fs.writeFileSync(entry, "import { context } from 'shared-context'; import { context as provided } from 'legacy-provider'; export const sameContext = context === provided;");
    const output = path.join(directory, 'output');
    const stats = await new Promise((resolve, reject) => {
      const compiler = webpack(buildRuntimeConfig(directory, entry, output, true));
      compiler.run((error, result) => compiler.close(closeError => error || closeError ? reject(error || closeError) : resolve(result)));
    });
    assert.equal(stats.hasErrors(), false, stats.toString({ all: false, errors: true }));
    assert.equal(require(path.join(output, 'entry.js')).sameContext, true, 'browser module graph must not split conditional contexts');
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

test('UI runtime supplies genuine WebCrypto to the JSDOM browser boundary', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'runtime-crypto-'));
  try {
    fs.writeFileSync(path.join(directory, 'package.json'), JSON.stringify({ name: 'runtime-crypto' }));
    fs.mkdirSync(path.join(directory, 'node_modules'), { recursive: true });
    for (const name of ['react', 'react-dom']) fs.symlinkSync(path.dirname(require.resolve(name + '/package.json')), path.join(directory, 'node_modules', name));
    const fixture = path.join(directory, 'probe.cjs');
    fs.writeFileSync(fixture, `exports.run = async ({ assert }) => {
      assert.equal(window.crypto, require('node:crypto').webcrypto);
      const bytes = new Uint8Array(16);
      assert.equal(window.crypto.getRandomValues(bytes), bytes);
      assert.equal(bytes.byteLength, 16);
      return { assertions: ['actual browser WebCrypto API is available'] };
    };`);
    assert.equal(executeProbe(fixture, 'form', directory).exit, 0);
    assert.equal(executeProbe(fixture, 'data-grid', directory).exit, 0);
    assert.equal(executeProbe(fixture, 'stage', directory).exit, 0);
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

test('service runtime loads a real asynchronous chunk in Node without a browser script URL', async () => {
  const webpack = require('webpack');
  const { buildRuntimeConfig } = require('../scripts/verify-production-runtime.cjs');
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'runtime-service-chunk-'));
  try {
    const entry = path.join(directory, 'entry.js');
    fs.writeFileSync(entry, "export async function loadBatch() { return (await import('./batch.js')).value; }");
    fs.writeFileSync(path.join(directory, 'batch.js'), 'export const value = 42;');
    const output = path.join(directory, 'output');
    const stats = await new Promise((resolve, reject) => {
      const compiler = webpack(buildRuntimeConfig(directory, entry, output, false));
      compiler.run((error, result) => compiler.close(closeError => error || closeError ? reject(error || closeError) : resolve(result)));
    });
    assert.equal(stats.hasErrors(), false, stats.toString({ all: false, errors: true }));
    assert.equal(await require(path.join(output, 'entry.js')).loadBatch(), 42);
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});

test('a fresh process executes emitted fixture bytes and reports HTTP observations', () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'runtime-process-unit-'));
  try {
    const fixture=path.join(dir,'probe.cjs');
    fs.writeFileSync(fixture,`exports.run=async send=>{const r=await send('https://tenant.test/sites/standalone/_api/web',{method:'GET',headers:{'X-Standalone':'context'}});const t=await send('https://tenant.test/sites/other/_api/web',{method:'GET',headers:{'X-Standalone':'targeted'}});return {web:await r.json(),targeted:await t.json(),resolved:'https://tenant.test/sites/other'};};`);
    const first=executeProbe(fixture,'context',dir);
    const second=executeProbe(fixture,'context',dir);
    assert.equal(first.scenario,'context');assert.equal(first.calls.length,2);
    assert.notEqual(first.pid,second.pid);assert.notEqual(first.pid,process.pid);
    assert.equal(first.exit,0);assert.equal(first.signal,null);assert.deepEqual(first.failureMarkers,[]);
  } finally {fs.rmSync(dir,{recursive:true,force:true});}
});
test('a small mutation fixture records its intended assertion failure', () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'runtime-mutation-unit-'));
  try {
    const fixture=path.join(dir,'probe.cjs');fs.writeFileSync(fixture,"exports.run=async()=>{throw Error('lost-registration-unit');};");
    assert.throws(()=>executeProbe(fixture,'context',dir),error=>error.exit===1 && error.stderr.includes('lost-registration-unit') && error.failureMarkers.length>0);
  } finally {fs.rmSync(dir,{recursive:true,force:true});}
});
test('matrix awaits work, writes progress without claiming PASS and aggregates independent failures', async () => {
  const report={complete:false,passed:false,probes:[],warnings:[],failures:[]};const snapshots=[];
  await runRuntimeMatrix({report,cases:[{scenario:'first',variant:'root'},{scenario:'second',variant:'leaf'}],compile:async probe=>({...probe,warnings:[],bundlePath:'fixture'}),execute:async (_,scenario)=>{await Promise.resolve();if(scenario==='first')throw Object.assign(Error('assertion failed'),{exit:1,signal:null,failureMarkers:['assertion failed']});return {exit:0,assertions:['second works']};},save:()=>snapshots.push({passed:report.passed,complete:report.complete,count:report.probes.length})});
  assert.equal(report.complete,true);assert.equal(report.passed,false);assert.equal(report.probes.length,2);
  assert.equal(report.failures[0].exit,1);assert.equal(report.probes[1].status,'passed');
  assert.ok(snapshots.slice(0,-1).every(s=>!s.complete&&!s.passed));assert.equal(snapshots.at(-1).count,2);
});
test('a real Webpack alias to a toolkit package at an arbitrary path is rejected', async () => {
  const webpack=require('webpack');
  const {statsOptions}=require('../scripts/bundle-contract.internal.cjs');
  const {assertToolkitOrigin}=require('../scripts/verify-production-runtime.cjs');
  const directory=fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(),'origin-unit-')));
  try {
    const copy=path.join(directory,'arbitrary'),consumer=path.join(directory,'consumer');
    fs.mkdirSync(path.join(copy,'lib'),{recursive:true});fs.mkdirSync(consumer);
    fs.writeFileSync(path.join(copy,'package.json'),JSON.stringify({name:'@apvee/spfx-react-toolkit',main:'lib/index.js'}));
    fs.writeFileSync(path.join(copy,'lib/index.js'),'export const value = 42;');
    fs.writeFileSync(path.join(consumer,'entry.js'),"import {value} from '@apvee/spfx-react-toolkit'; export const result=value;");
    const stats=await new Promise((resolve,reject)=>{const compiler=webpack({mode:'production',context:consumer,entry:'./entry.js',output:{path:path.join(directory,'dist'),library:{type:'commonjs2'}},resolve:{alias:{'@apvee/spfx-react-toolkit':copy}},optimization:{concatenateModules:true}});compiler.run((error,result)=>compiler.close(closeError=>error||closeError?reject(error||closeError):resolve(result)));});
    assert.equal(stats.hasErrors(),false);
    assert.throws(()=>assertToolkitOrigin(stats.toJson(statsOptions()),consumer),/Toolkit origin escaped consumer/);
    fs.mkdirSync(path.join(consumer,'node_modules/@apvee'),{recursive:true});
    fs.cpSync(copy,path.join(consumer,'node_modules/@apvee/spfx-react-toolkit'),{recursive:true});
    assert.doesNotThrow(()=>assertToolkitOrigin({chunks:[{id:1,modules:[{nameForCondition:path.join(consumer,'node_modules/@apvee/spfx-react-toolkit/lib/index.js')}]}]},consumer));
  } finally {fs.rmSync(directory,{recursive:true,force:true});}
});

test('canonical runtime matrix includes all 51 installed probes and all three trim aliases', () => {
  const { productionRuntimeCases } = require('../scripts/verify-production-runtime.cjs');
  const cases = productionRuntimeCases();
  assert.equal(cases.length, 51);
  assert.equal(new Set(cases.map(item => `${item.scenario}/${item.variant}`)).size, 51);
  for (const family of ['sharepoint-permission-trim', 'sharepoint-group-trim', 'entra-group-trim']) {
    assert.deepEqual(cases.filter(item => item.scenario === family).map(item => item.variant), ['root', 'domain', 'control']);
    const template = fs.readFileSync(path.resolve(__dirname, `fixtures/tree-shaking/runtime/${family}.tsx`), 'utf8');
    assert.match(template, /from '__TOOLKIT__'/); assert.match(template, /SPFxWebPartProvider/); assert.match(template, /runTrimProbe/);
  }
});
