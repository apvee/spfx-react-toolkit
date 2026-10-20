const test=require('node:test'),assert=require('node:assert/strict');
const {createHarness}=require('./services-test-harness.cjs'); const h=createHarness();
const protocol=h.load('controls/file-folder-picker/file-folder-picker-protocol.internal.ts');
test('authentication rejects malformed/foreign URL audiences and challenges without reflecting payloads',()=>{
 for(const resource of ['https://evil.sharepoint.com','https://tenant.sharepoint.com@evil.test','http://tenant.sharepoint.com','https://tenant.sharepoint.com/path','https://tenant.sharepoint.com/?x=y']) assert.throws(()=>protocol.readFilePickerTokenRequest({resource},['https://tenant.sharepoint.com']));
 for(const field of ['claims','claimsChallenge']) assert.throws(()=>protocol.readFilePickerTokenRequest({resource:'https://tenant.sharepoint.com',[field]:'private'},['https://tenant.sharepoint.com']),{code:'auth-challenge'});
 assert.deepEqual(protocol.readFilePickerTokenRequest({resource:'https://tenant.sharepoint.com/'},['https://tenant.sharepoint.com']),{resource:'https://tenant.sharepoint.com',type:undefined});
 let reads=0;const message={type:'command',id:'1'};Object.defineProperty(message,'data',{get(){reads++;return {command:'authenticate'}}});assert.equal(protocol.readFilePickerCommand(message),undefined);assert.equal(reads,0);
});
test('hosted navigation POST uses exact page, encoded configuration and ephemeral hidden token form',()=>{
 const {postFileFolderPicker}=h.load('controls/file-folder-picker/file-folder-picker-navigation.internal.ts');
 let submitted;
 const submit=window.HTMLFormElement.prototype.submit;
 window.HTMLFormElement.prototype.submit=function(){submitted={method:this.method,action:this.action,target:this.target,hidden:this.hidden,token:this.elements.namedItem('access_token').value};};
 try { postFileFolderPicker(document,'picker-one','https://tenant.sharepoint.com/sites/hr',{sdk:'8.0',messaging:{channelId:'one',origin:'https://host.test'}},'secret','it-it'); }
 finally {window.HTMLFormElement.prototype.submit=submit;}
 assert.equal(submitted.method,'post');assert.equal(submitted.target,'picker-one');assert.equal(submitted.hidden,true);assert.equal(submitted.token,'secret');
 const url=new URL(submitted.action);assert.equal(url.pathname,'/sites/hr/_layouts/15/FilePicker.aspx');assert.equal(url.searchParams.get('locale'),'it-it');assert.equal(JSON.parse(url.searchParams.get('filePicker')).messaging.channelId,'one');assert.equal(url.href.includes('secret'),false);assert.equal(document.querySelector('[name="access_token"]'),null);
});
test('observed SharePoint_SelfIssued hint still requires exact SharePoint resource and rejects Graph',()=>{
 const request=protocol.readFilePickerTokenRequest({type:'SharePoint_SelfIssued',resource:'https://tenant.sharepoint.com'},['https://tenant.sharepoint.com']);assert.equal(request.resource,'https://tenant.sharepoint.com');assert.equal(request.type,'SharePoint_SelfIssued');
 for(const resource of ['https://graph.microsoft.com','https://evil.sharepoint.com'])assert.throws(()=>protocol.readFilePickerTokenRequest({type:'SharePoint_SelfIssued',resource},['https://tenant.sharepoint.com']));
 assert.throws(()=>protocol.readFilePickerTokenRequest({type:'Graph',resource:'https://tenant.sharepoint.com'},['https://tenant.sharepoint.com']));
});

test('observed MSAL hint uses only the exact configured audience', () => {
 assert.equal(protocol.readFilePickerTokenRequest({type:'SharePoint_MSAL',resource:'https://tenant.sharepoint.com'},['https://tenant.sharepoint.com']).resource,'https://tenant.sharepoint.com');
 assert.throws(()=>protocol.readFilePickerTokenRequest({type:'SharePoint_MSAL',resource:'https://tenant-my.sharepoint.com'},['https://tenant.sharepoint.com']),{code:'auth-resource'});
 assert.throws(()=>protocol.readFilePickerTokenRequest({type:'Graph',resource:'https://graph.microsoft.com'},['https://tenant.sharepoint.com']),{code:'auth-resource'});
});

test('Graph requires explicit opt-in and exact audience/type; claims require validated nested payload', () => {
 const allowed=['https://tenant.sharepoint.com'];
 const options={graphEnabled:true,claimsChallengeEnabled:true};
 assert.deepEqual(protocol.readFilePickerTokenRequest({resource:'https://graph.microsoft.com',type:'Graph'},allowed,options),{resource:'https://graph.microsoft.com',type:'Graph'});
 assert.throws(()=>protocol.readFilePickerTokenRequest({resource:'https://graph.microsoft.com',type:'SharePoint'},allowed,options));
 assert.throws(()=>protocol.readFilePickerTokenRequest({resource:'https://tenant.sharepoint.com',type:'Graph'},allowed,options));
 const claims='{"access_token":{"xms_cc":{"values":["cp1"]}}}';
 assert.equal(protocol.readFilePickerTokenRequest({resource:allowed[0],claimsChallenge:{claims}},allowed,options).claims,claims);
 for(const challenge of [{claims:''},{claims:12},{claims:'not-json'}]) assert.throws(()=>protocol.readFilePickerTokenRequest({resource:allowed[0],claimsChallenge:challenge},allowed,options));
 assert.throws(()=>protocol.readFilePickerTokenRequest({resource:allowed[0],claimsChallenge:{claims}},allowed,{graphEnabled:true}));
});

test('authentication diagnostics separate URL shape from audience without exposing input', () => {
 const events=[]; const allowed=['https://tenant.sharepoint.com'];
 const options={onDiagnostic:event=>events.push(event)};
 for(const [data,stage] of [
  [{},'resource-missing'],[{resource:'private not a URL'},'resource-invalid-url'],
  [{resource:allowed[0]+'/private?secret=claims'},'resource-invalid-shape'],
  [{resource:'https://tenant-my.sharepoint.com',type:'SharePoint_MSAL'},'audience-denied'],
  [{resource:allowed[0],type:'PRIVATE-TYPE'},'type-unsupported'],
  [{resource:allowed[0],type:'Graph'},'type-audience-mismatch'],
  [{resource:allowed[0],claims:'PRIVATE-CLAIMS'},'claims-denied'],
 ]) { assert.throws(()=>protocol.readFilePickerTokenRequest(data,allowed,options)); assert.equal(events.at(-1).stage,stage); }
 assert.equal(events[2].resource,'source'); assert.equal(events[2].nonRootPath,true);
 assert.equal(events[3].resource,'tenant-personal-candidate');
 assert.equal(events[4].type,'unsupported');
 assert.doesNotMatch(JSON.stringify(events),/private|secret|claims=|PRIVATE|https:|sharepoint\.com/);
 assert.equal(protocol.readFilePickerTokenRequest({resource:allowed[0]},allowed,{onDiagnostic(){throw new Error('observer');}}).resource,allowed[0]);
});

test('debug diagnostics use the owner document exact opt-in and never enrich with user context', () => {
 const diagnostics=h.load('controls/file-folder-picker/file-folder-picker-diagnostics.internal.ts');
 for(const search of ['','?debug=false','?debug=1','?other=debug=true']) assert.equal(diagnostics.filePickerDebugObserver({location:{search}}),undefined);
 const messages=[]; const previous=console.debug; console.debug=(...args)=>messages.push(args);
 try {
  diagnostics.filePickerDebugObserver({location:{search:'?debug=true'}})({stage:'audience-denied',resource:'other',type:'missing'});
  assert.equal(messages.length,1); assert.equal(messages[0][0],'[FileFolderPicker authentication]');
  assert.deepEqual(JSON.parse(messages[0][1]),{stage:'audience-denied',resource:'other',type:'missing'});
 } finally {console.debug=previous;}
});
