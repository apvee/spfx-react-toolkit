const test = require('node:test');
const assert = require('node:assert/strict');
const { collectEmittedModules, collectReachableAssets, finalizeReport, renderConsumer, dataCalls, retentionObservations } = require('../scripts/bundle-contract.internal.cjs');
function fullEnforceReport() {
  const contract=require('./fixtures/tree-shaking/contract.json');
  const {checkComparisons}=require('../scripts/bundle-contract.internal.cjs');
  const modules={'settings-group':'@apvee/spfx-react-toolkit/lib/controls/settings-group/settings-group.js','accordion-group':'@apvee/spfx-react-toolkit/lib/controls/accordion-group/accordion-group.js','fluent-card':'@fluentui/react-card/lib/components/Card/Card.js','fluent-accordion':'@fluentui/react-accordion/lib/components/Accordion/Accordion.js',pnp:'@pnp/sp/webs/index.js','styles-engine':'@apvee/spfx-react-toolkit/lib/helpers/styles/resolve.internal.js','list-registration':'@pnp/sp/lists/index.js','search-registration':'@pnp/sp/search/index.js','batch-registration':'@pnp/sp/batching.js','formatted-text':'@apvee/spfx-react-toolkit/lib/controls/formatted-text/formatted-text.js','markdown':'unified/index.js','fluent-link':'@fluentui/react-link/lib/index.js','typography':'@apvee/spfx-react-toolkit/lib/helpers/styles/typography.js','relative-time':'@apvee/spfx-react-toolkit/lib/controls/relative-time/relative-time.js','fluent-tooltip':'@fluentui/react-tooltip/lib/components/Tooltip/Tooltip.js','iana-lookup':'@apvee/spfx-react-toolkit/lib/helpers/spfx-time-zone.internal.js','auto-skeleton':'@apvee/spfx-react-toolkit/lib/controls/auto-skeleton/auto-skeleton.js','state-message':'@apvee/spfx-react-toolkit/lib/controls/state-message/state-message.js','state-artwork':'@apvee/spfx-react-toolkit/lib/controls/state-message/illustrations/default-illustration.js','fluent-button':'@fluentui/react-button/lib/components/Button/Button.js','fluent-skeleton':'@fluentui/react-skeleton/lib/components/Skeleton/Skeleton.js',form:'@apvee/spfx-react-toolkit/lib/controls/form/form.js','react-hook-form':'react-hook-form/dist/index.esm.mjs','form-resolver':'@hookform/resolvers/standard-schema/dist/standard-schema.mjs','form-schema':'@standard-schema/utils/dist/index.js','form-date':'@fluentui/react-datepicker-compat/lib/components/DatePicker/DatePicker.js','form-time':'@fluentui/react-timepicker-compat/lib/components/TimePicker/TimePicker.js'};
  Object.assign(modules, { 'data-grid': '@apvee/spfx-react-toolkit/lib/controls/data-grid/data-grid.js', charts: '@fluentui/react-charts/lib/components/Sparkline/Sparkline.js', d3: 'd3-shape/src/line.js', shell: '@apvee/spfx-react-toolkit/lib/controls/web-part-shell/web-part-shell.js', editor: '@apvee/spfx-react-toolkit/lib/controls/web-part-shell/title/web-part-title-editor.internal.js', tiptap: '@tiptap/core/dist/index.js', prosemirror: 'prosemirror-state/dist/index.js' });
  Object.assign(modules, { stage: '@apvee/spfx-react-toolkit/lib/controls/stage/stage.js', 'fluent-dialog': '@fluentui/react-dialog/lib/components/Dialog/Dialog.js', 'fluent-drawer': '@fluentui/react-drawer/lib/components/OverlayDrawer/OverlayDrawer.js', 'fluent-popover': '@fluentui/react-popover/lib/components/Popover/Popover.js', griffel: '@griffel/react/lib/makeStyles.js', 'style-descriptors': '@apvee/spfx-react-toolkit/lib/helpers/styles/width.js', 'styles-compiler': '@apvee/spfx-react-toolkit/lib/helpers/styles/resolve.internal.js' });
  const fixtures=contract.fixtures.flatMap(fixture=>Object.keys(fixture.variants).map(variant=>({id:`${fixture.id}/${variant}`,status:'measured',initialChunkIds:[1],emittedModules:fixture.requiredModuleGroups.map(group=>({resource:modules[group],chunks:fixture.forbiddenInitialModuleGroups?.includes(group)?[2]:[1]})),retentionObservations:[],totals:{union:{gzip:100}},fixtureSha256:'fixture',configSha256:'config',assets:[]})));
  return {schemaVersion:1,mode:'enforce',complete:true,passed:true,provenance:{tarballSha256:'abc'},contractSha256:require('node:crypto').createHash('sha256').update(JSON.stringify(contract)).digest('hex'),fixtures,comparisons:checkComparisons(fixtures,contract).comparisons,failures:[],errors:[]};
}
test('concatenated emitted modules are traversed', () => {
  const stats = { chunks: [{ id: 1, modules: [{ name: './entry + 2 modules', modules: [{ name: './node_modules/@pnp/sp/webs/index.js' }] }] }] };
  assert.ok(collectEmittedModules(stats).some(m => m.resource === '@pnp/sp/webs/index.js'));
});
test('unused parsed module is not emitted', () => {
  assert.deepEqual(collectEmittedModules({ modules: [{ name: './unused.js', chunks: [] }], chunks: [{ id: 1, modules: [{ name: './used.js' }] }] }).map(m => m.resource), ['used.js']);
});
test('shared lazy asset is counted once', () => {
  const stats = { entrypoints: { main: { chunks: [1] } }, chunks: [{ id: 1, files: ['main.js'], children: [2, 3] }, { id: 2, files: ['shared.js'], children: [] }, { id: 3, files: ['shared.js'], children: [] }] };
  assert.deepEqual(collectReachableAssets(stats), { initial: ['main.js'], async: ['shared.js'], union: ['main.js', 'shared.js'] });
  assert.throws(() => collectReachableAssets({ ...stats, chunks: stats.chunks.slice(0, 1) }), /Missing reachable chunk/);
});
test('audit report cannot satisfy enforce', () => {
  assert.equal(finalizeReport({ mode: 'audit', complete: true, failures: [] }).passed, null);
  assert.equal(finalizeReport({ mode: 'enforce', complete: true, failures: [] }).passed, false);
  assert.equal(finalizeReport(fullEnforceReport()).passed,true);
  assert.equal(finalizeReport({ mode: 'enforce', complete: true, failures: ['missing'] }).passed, false);
});
test('consumer slots are strict and width pairs are coherent', () => {
  assert.equal(renderConsumer("import { __WIDTH_EXPORT__ } from '__TOOLKIT__'; export const value = __WIDTH_FULL__;", { imports: { TOOLKIT: 'pkg' }, widthExport: 'width', widthExpression: 'width.full' }), "import { width } from 'pkg'; export const value = width.full;");
  assert.throws(() => renderConsumer('__HOOK__', { imports: {} }), /slot/);
  assert.throws(() => renderConsumer('__WIDTH_FULL__', { imports: {}, widthExport: 'full', widthExpression: 'width.full' }), /width/);
});
test('retention parser observes static descriptor constructor calls', () => {
  assert.deepEqual(retentionObservations('x("width", "100%", {fallback:false}); y("body1", []);'), [{ kind: 'declaration', property: 'width', staticValue: '100%' }, { kind: 'recipe', name: 'body1' }]);
});

test('legacy retention parser preserves value observation', () => {
  assert.deepEqual(dataCalls('x("width", "100%", {fallback:false});'), [{kind:'declaration', property:'width', value:'100%'}]);
});
test('contract rejects unknown module groups and malformed comparisons', () => {
  const {validateContract} = require('../scripts/bundle-contract.internal.cjs');
  assert.throws(() => validateContract({schemaVersion:1,moduleGroups:{},fixtures:[{id:'helper',template:'helper.ts',variants:{root:{imports:{}}},forbiddenModuleGroups:['unknown'],requiredModuleGroups:[]}],comparisons:[]}), /Unknown module group/);
});
test('DataGrid and factory aliases enforce complete-family cost and exclude query/Form/PnP code', () => {
  const contract = require('./fixtures/tree-shaking/contract.json');
  for (const id of ['data-grid', 'data-grid-factory']) {
    const fixture = contract.fixtures.find(item => item.id === id);
    assert.ok(fixture, `${id}: missing optimized consumer`);
    assert.deepEqual(Object.keys(fixture.variants), ['root', 'domain', 'control']);
    for (const group of ['data-grid', 'charts', 'd3']) assert.ok(fixture.requiredModuleGroups.includes(group), `${id}: ${group}`);
    for (const group of ['grid-query', 'pnp', 'spfx-runtime', 'provider', 'form', 'react-hook-form', 'form-resolver', 'form-schema']) {
      assert.ok(fixture.forbiddenModuleGroups.includes(group), `${id}: ${group}`);
    }
    for (const alias of ['root', 'domain']) {
      const comparison = contract.comparisons.find(item => item.candidate === `${id}/${alias}`);
      assert.equal(comparison?.reference, `${id}/control`);
      assert.equal(comparison.maxIncreaseBytes, 0, 'additive aliases must not retain unrelated features');
    }
  }
  for (const fixture of contract.fixtures.filter(item => !['data-grid', 'data-grid-factory'].includes(item.id))) {
    for (const group of ['data-grid', 'charts', 'd3']) assert.ok(fixture.forbiddenModuleGroups.includes(group), `${fixture.id}: ${group}`);
  }
});
test('unrelated consumers reject actual grid/chart/D3 module emissions independently of size', () => {
  const contract = require('./fixtures/tree-shaking/contract.json');
  const { checkFixture } = require('../scripts/bundle-contract.internal.cjs');
  const fixture = contract.fixtures.find(item => item.id === 'helper');
  for (const resource of [
    '@apvee/spfx-react-toolkit/lib/controls/data-grid/data-grid.js',
    '@fluentui/react-table/lib/components/DataGrid/DataGrid.js',
    '@fluentui/react-charts/lib/components/Sparkline/Sparkline.js',
    '@fluentui/chart-utilities/lib/utilities/index.js',
    'd3-scale/src/index.js',
    'd3-shape/src/line.js',
  ]) {
    const failures = checkFixture({ id: 'helper/root', emittedModules: [{ resource }], retentionObservations: [] }, fixture, contract);
    assert.ok(failures.some(message => /forbidden module group/.test(message)), resource);
  }
});
test('descriptor and emitted group checks produce concrete failures', () => {
  const {checkFixture} = require('../scripts/bundle-contract.internal.cjs');
  const contract = {moduleGroups:{pnp:['^@pnp/'],engine:['^toolkit/engine$']}};
  const fixture = {forbiddenModuleGroups:['pnp'],requiredModuleGroups:['engine'],retainedDescriptors:{recipeNames:[],declarationProperties:['width'],staticValues:{width:['100%']},maxCallCount:1}};
  const failures = checkFixture({id:'width/root',emittedModules:[{resource:'@pnp/sp/webs/index.js'}],retentionObservations:[{kind:'recipe',name:'body2'},{kind:'declaration',property:'width',staticValue:'auto'}]},fixture,contract);
  assert.equal(failures.length,5);
});
test('truncated nested stats are rejected instead of hiding emitted dependencies', () => {
  assert.throws(() => collectEmittedModules({chunks:[{id:1,modules:[{name:'entry',modules:[{type:'dependent modules',filteredChildren:12}]}]}]}), /truncated/i);
});
test('nonliteral descriptor values are not reported as static values', () => {
  assert.deepEqual(retentionObservations('x("fontSize", theme.body1.fontSize, {fallback:"inherit"});'), [{kind:'declaration',property:'fontSize'}]);
});
test('real production concatenation exposes emitted leaves without parsed orphan inventory', async () => {
  const fs = require('node:fs');
  const os = require('node:os');
  const path = require('node:path');
  const webpack = require('webpack');
  const {statsOptions} = require('../scripts/bundle-contract.internal.cjs');
  const dir = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(),'tree-shaking-stats-test-')));
  try {
    fs.writeFileSync(path.join(dir,'entry.js'),'import {value} from "./leaf.js"; import {unused} from "./unused.js"; export const result = value;');
    fs.writeFileSync(path.join(dir,'leaf.js'),'export const value = 42;');
    fs.writeFileSync(path.join(dir,'unused.js'),'export const unused = 99;');
    const stats = await new Promise((resolve,reject) => {
      const compiler = webpack({context:dir,mode:'production',cache:false,entry:'./entry.js',output:{path:path.join(dir,'dist'),library:{type:'commonjs2'}},optimization:{usedExports:true,concatenateModules:true,minimize:true}});
      compiler.run((error,result)=>compiler.close(closeError=>error||closeError?reject(error||closeError):resolve(result)));
    });
    assert.equal(stats.hasErrors(),false);
    const detail=stats.toJson(statsOptions());
    assert.equal(detail.modules,undefined);
    assert.ok(collectEmittedModules(detail,dir).some(m=>m.resource==='leaf.js'));
  } finally {fs.rmSync(dir,{recursive:true,force:true});}
});
test('replay comparison requires identical asset bytes hashes and stable provenance', () => {
  const {compareBundleReports} = require('../scripts/bundle-contract.internal.cjs');
  const report = {schemaVersion:1,mode:'audit',provenance:{consumerRoot:'/first',toolkitDir:'/first/node_modules/toolkit',tarballSha256:'abc',versions:{webpack:'5'}},fixtures:[{id:'helper/root',status:'measured',fixtureSha256:'fixture',configSha256:'config',assets:[{name:'entry.js',raw:100,gzip:50,brotli:40,sha256:'bytes',initial:true}],totals:{union:{raw:100,gzip:50,brotli:40}}}]};
  const relocated = JSON.parse(JSON.stringify(report)); relocated.provenance.consumerRoot='/second';relocated.provenance.toolkitDir='/second/node_modules/toolkit';
  assert.deepEqual(compareBundleReports(report,relocated),[]);
  relocated.fixtures[0].assets[0].sha256='drift';
  assert.match(compareBundleReports(report,relocated)[0],/asset drift/);
  relocated.provenance.tarballSha256='different';
  assert.ok(compareBundleReports(report,relocated).some(f=>/provenance/.test(f)));
});
test('replay comparison rejects missing fixtures unless an explicit subset was selected', () => {
  const {compareBundleReports}=require('../scripts/bundle-contract.internal.cjs');
  const fixture=id=>({id,status:'unavailable'});
  const reference={provenance:{},fixtures:[fixture('helper/domain'),fixture('provider/domain')]};
  const subset={provenance:{},fixtures:[fixture('helper/domain')]};
  assert.ok(compareBundleReports(reference,subset).some(f=>/missing fixture/.test(f)));
  assert.deepEqual(compareBundleReports(reference,subset,{fixtureIds:['helper']}),[]);
});
test('measurement runtime provenance records Node and npm versions', () => {
  const {readRuntimeToolchain}=require('../scripts/bundle-contract.internal.cjs');
  const versions=readRuntimeToolchain();
  assert.equal(versions.node,process.version);
  assert.match(versions.npm,/^\d+\.\d+\.\d+$/);
});
test('initial and progress enforce snapshots never claim a completed gate', () => {
  const initial={schemaVersion:1,mode:'enforce',complete:false,fixtures:[],comparisons:[],warnings:[],errors:[],failures:[]};
  assert.equal(finalizeReport(initial).passed,false);
  assert.equal(finalizeReport({mode:'enforce',failures:[],errors:[]}).passed,false);
  assert.equal(finalizeReport({...initial,fixtures:[{id:'helper/root',status:'measured'}]}).passed,false);
  assert.equal(finalizeReport({...initial,complete:true}).passed,false);
  assert.equal(finalizeReport({...initial,complete:true,failures:['missing comparison']}).passed,false);
  assert.equal(finalizeReport({...initial,complete:true,errors:['build error']}).passed,false);
});

test('isolation contract rejects upstream typography even with no retained descriptor calls', () => {
  const {checkFixture} = require('../scripts/bundle-contract.internal.cjs');
  const contract = require('./fixtures/tree-shaking/contract.json');
  const width = contract.fixtures.find(f => f.id === 'sx-width-only');
  const failures = checkFixture({id:'sx-width-only/root',emittedModules:[{resource:'@fluentui/react-theme/lib/tokens.js'}],retentionObservations:[]},width,contract);
  assert.ok(failures.some(f => f.includes('typography')));
  assert.ok(contract.fixtures.find(f => f.id === 'helper').forbiddenModuleGroups.includes('griffel'));
  assert.ok(contract.fixtures.find(f => f.id === 'stable-callback').forbiddenModuleGroups.includes('provider'));
});


for (const [budget, boundary] of [[1023, 1023], [512, 512], [0, 0]]) {
  test(`gzip budget ${budget} accepts the exact boundary and rejects one byte over`, () => {
    const { checkComparisons } = require('../scripts/bundle-contract.internal.cjs');
    const contract = {comparisons:[{candidate:'helper/root',reference:'helper/leaf',metric:'gzip',maxIncreaseBytes:budget}]};
    const fixture = (id, gzip) => ({id,status:'measured',totals:{union:{gzip}}});
    assert.deepEqual(checkComparisons([fixture('helper/root',100+boundary),fixture('helper/leaf',100)],contract).failures,[]);
    assert.match(checkComparisons([fixture('helper/root',101+boundary),fixture('helper/leaf',100)],contract).failures[0],/Gzip comparison/);
  });
}
test('informational engine comparisons report a delta without failing enforce', () => {
  const {checkComparisons}=require('../scripts/bundle-contract.internal.cjs');
  const result=checkComparisons([{id:'sx/leaf',status:'measured',totals:{union:{gzip:1900}}},{id:'direct/control',status:'measured',totals:{union:{gzip:100}}}],{comparisons:[{candidate:'sx/leaf',reference:'direct/control',metric:'gzip',informational:true}]});
  assert.deepEqual(result.failures,[]);assert.equal(result.comparisons[0].increaseBytes,1800);
});
test('forbidden emission fails independently of a negative byte delta', () => {
  const {checkFixture,checkComparisons}=require('../scripts/bundle-contract.internal.cjs');
  const contract={moduleGroups:{pnp:['^@pnp/']},comparisons:[{candidate:'helper/root',reference:'helper/leaf',metric:'gzip',maxIncreaseBytes:512}]};
  const item={id:'helper/root',status:'measured',totals:{union:{gzip:10}},emittedModules:[{resource:'@pnp/sp/webs/index.js'}],retentionObservations:[]};
  assert.equal(checkFixture(item,{forbiddenModuleGroups:['pnp'],requiredModuleGroups:[]},contract).length,1);
  assert.deepEqual(checkComparisons([item,{id:'helper/leaf',status:'measured',totals:{union:{gzip:100}}}],contract).failures,[]);
});
test('baseline writing accepts only a complete enforce report and preserves provenance', () => {
  const {baselineSnapshot}=require('../scripts/bundle-contract.internal.cjs');
  assert.throws(()=>baselineSnapshot({mode:'audit',passed:null,complete:true}),/complete enforce PASS/);
  assert.throws(()=>baselineSnapshot({mode:'enforce',passed:false,complete:false}),/complete enforce PASS/);
  const report=fullEnforceReport();
  assert.deepEqual(baselineSnapshot(report).provenance,report.provenance);
  assert.equal(baselineSnapshot(report).fixtures[0].totals.union.gzip,100);
});
test('CLI defaults to enforce and makes attribution a separate diagnostic', () => {
  const {parseArgs}=require('../scripts/verify-bundle-contract.cjs');
  assert.deepEqual(parseArgs(['--consumer-root','/tmp/consumer']).mode,'enforce');
  assert.equal(parseArgs(['--consumer-root','/tmp/consumer','--attribution']).attribution,true);
  assert.throws(()=>parseArgs(['--consumer-root','/tmp/consumer','--mode','audit','--write-baseline','baseline.json']),/enforce/);
});

test('installed gate orchestration awaits every verifier, rejects audit and preserves independent checks', async () => {
  const {runInstalledGates}=require('../scripts/bundle-contract.internal.cjs');
  const events=[],failures=[];
  await runInstalledGates({consumerRoot:'/consumer',outputDir:'/evidence',failures,
    bundle:async()=>{await Promise.resolve();events.push('bundle');return {mode:'audit',complete:true,passed:null};},
    runtime:async()=>{events.push('runtime');throw Error('runtime failed');},
    mutations:async()=>{events.push('mutations');return {complete:true,passed:true};},
    independent:async()=>events.push('independent')});
  assert.deepEqual(events,['bundle','runtime','mutations','independent']);
  assert.equal(failures.length,2);assert.ok(failures[0].message.includes('complete enforce PASS'));assert.ok(failures[1].message.includes('runtime failed'));
});

test('baseline completeness independently requires canonical contract, all measured fixture IDs and comparisons', () => {
  const {baselineSnapshot}=require('../scripts/bundle-contract.internal.cjs');
  const variants=[report=>{report.fixtures=[];},report=>{report.fixtures.pop();},report=>{report.fixtures[0].status='unavailable';},report=>{report.fixtures[0].id='replacement/root';},report=>{report.comparisons=[];},report=>{report.comparisons[0].maxIncreaseBytes=999999;},report=>{report.contractSha256='altered';},report=>{report.fixtures[0].emittedModules.push({resource:'@pnp/sp/webs/index.js'});}];
  for(const mutate of variants){const report=fullEnforceReport();mutate(report);assert.equal(finalizeReport(report).passed,false);assert.throws(()=>baselineSnapshot(report),/canonical complete enforce PASS/);}
  assert.equal(baselineSnapshot(fullEnforceReport()).fixtures.length,81);
});
test('real CLI refuses empty, subset, altered-budget and weakened forbidden-group contracts before baseline write', () => {
  const fs=require('node:fs'),path=require('node:path'),os=require('node:os');
  const {spawnSync}=require('node:child_process');
  const directory=fs.mkdtempSync(path.join(os.tmpdir(),'canonical-cli-unit-'));
  const canonical=require('./fixtures/tree-shaking/contract.json');
  const cases=[
    ['empty',()=>({schemaVersion:1,moduleGroups:{},fixtures:[],comparisons:[]})],
    ['subset',contract=>{contract.fixtures=contract.fixtures.slice(0,1);contract.comparisons=[];return contract;}],
    ['budget',contract=>{contract.comparisons[0].maxIncreaseBytes=999999;return contract;}],
    ['forbidden',contract=>{contract.fixtures[0].forbiddenModuleGroups=[];return contract;}],
  ];
  try {
    for(const [name,alter] of cases){
      const file=path.join(directory,name+'.json'),baseline=path.join(directory,name+'-baseline.json');
      fs.writeFileSync(file,JSON.stringify(alter(structuredClone(canonical))));
      const result=spawnSync(process.execPath,[path.resolve(__dirname,'../scripts/verify-bundle-contract.cjs'),'--consumer-root',path.join(directory,'missing-consumer'),'--contract',file,'--output',path.join(directory,name),'--write-baseline',baseline],{encoding:'utf8'});
      assert.equal(result.status,1,name);assert.match(result.stderr,/Enforce requires the canonical complete production contract/,name);assert.equal(fs.existsSync(baseline),false,name);
    }
    const copy=path.join(directory,'identical-copy.json');fs.writeFileSync(copy,JSON.stringify({comparisons:canonical.comparisons,moduleGroups:canonical.moduleGroups,fixtures:canonical.fixtures,schemaVersion:canonical.schemaVersion},null,4));
    const result=spawnSync(process.execPath,[path.resolve(__dirname,'../scripts/verify-bundle-contract.cjs'),'--consumer-root',path.join(directory,'missing-consumer'),'--contract',copy],{encoding:'utf8'});
    assert.equal(result.status,1);assert.match(result.stderr,/ENOENT/);assert.doesNotMatch(result.stderr,/canonical complete production contract/);
  } finally {fs.rmSync(directory,{recursive:true,force:true});}
});

test('Form and factory contracts measure the complete family and forbid it in every unrelated fixture', () => {
  const { checkFixture } = require('../scripts/bundle-contract.internal.cjs');
  const contract = require('./fixtures/tree-shaking/contract.json');
  const groups = ['form', 'react-hook-form', 'form-resolver', 'form-schema', 'form-date', 'form-time'];
  for (const fixture of contract.fixtures) {
    if (['form', 'form-factory'].includes(fixture.id)) {
      assert.deepEqual(Object.keys(fixture.variants), ['root', 'domain', 'control']);
      for (const group of ['form', 'react-hook-form', 'form-resolver', 'form-schema', 'form-date', 'form-time']) assert.ok(fixture.requiredModuleGroups.includes(group));
      for (const group of ['pnp', 'spfx-runtime', 'provider', 'formatted-text', 'settings-group', 'accordion-group']) assert.ok(fixture.forbiddenModuleGroups.includes(group));
    } else for (const group of groups) assert.ok(fixture.forbiddenModuleGroups.includes(group), `${fixture.id}: ${group}`);
  }
  const helper = contract.fixtures.find(f => f.id === 'helper');
  for (const resource of ['@apvee/spfx-react-toolkit/lib/controls/form/form.js', 'react-hook-form/dist/index.esm.mjs', '@hookform/resolvers/standard-schema/dist/standard-schema.mjs', '@standard-schema/spec/dist/index.js', '@fluentui/react-datepicker-compat/lib/index.js', '@fluentui/react-calendar-compat/lib/index.js', '@fluentui/react-timepicker-compat/lib/index.js']) {
    assert.ok(checkFixture({ id: 'helper/root', emittedModules: [{ resource }], retentionObservations: [] }, helper, contract).some(f => f.includes('forbidden module group')), resource);
  }
});

test('optimized consumer resolves nested peer packages through standard Node hierarchy', async () => {
  const fs = require('node:fs');
  const os = require('node:os');
  const path = require('node:path');
  const webpack = require('webpack');
  const { buildConfig } = require('../scripts/bundle-contract.internal.cjs');
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'nested-bundle-peer-'));
  try {
    const owner = path.join(directory, 'node_modules/owner');
    const nested = path.join(owner, 'node_modules/nested-peer');
    fs.mkdirSync(nested, { recursive: true });
    fs.writeFileSync(path.join(owner, 'package.json'), JSON.stringify({ name: 'owner', module: 'index.js' }));
    fs.writeFileSync(path.join(owner, 'index.js'), "export { value } from 'nested-peer';");
    fs.writeFileSync(path.join(nested, 'package.json'), JSON.stringify({ name: 'nested-peer', module: 'index.js' }));
    fs.writeFileSync(path.join(nested, 'index.js'), 'export const value = 42;');
    const entry = path.join(directory, 'entry.js');
    fs.writeFileSync(entry, "export { value } from 'owner';");
    const output = path.join(directory, 'output');
    const compiler = webpack(buildConfig(directory, entry, output, false));
    const stats = await new Promise((resolve, reject) => compiler.run((error, stats) => compiler.close(closeError => error || closeError ? reject(error || closeError) : resolve(stats))));
    assert.equal(stats.hasErrors(), false, stats.toString({ all: false, errors: true }));
    assert.equal(require(path.join(output, 'entry.js')).value, 42);
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});
