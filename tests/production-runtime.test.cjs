const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { executeProbe, runRuntimeMatrix } = require('../scripts/verify-production-runtime.cjs');

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
