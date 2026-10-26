const test=require('node:test'), assert=require('node:assert/strict'),path=require('node:path'),fs=require('node:fs');
const {createHarness}=require('./services-test-harness.cjs');
const base=path.resolve(__dirname,'../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest');
function loadPart(){const factory=(targetProperty,properties)=>({targetProperty,properties});class Base{constructor(){this.properties={};this.refreshes=0;this.context={pageContext:{web:{absoluteUrl:'https://tenant.sharepoint.com/sites/hr'}},propertyPane:{refresh:()=>this.refreshes++}}}}
 const h=createHarness({'@microsoft/sp-core-library':{Version:{parse:x=>x}},'@microsoft/sp-property-pane':Object.fromEntries(['TextField','Dropdown','Toggle','Label'].map(name=>['PropertyPane'+name,factory])),'@microsoft/sp-webpart-base':{BaseClientSideWebPart:Base},'@apvee/spfx-react-toolkit':{getWebPartShellPropertiesMetadata:()=>({})},'@apvee/spfx-react-toolkit/property-pane/pickers':Object.fromEntries(['Site','List','Field','View','ContentType','ListItem','Principal'].map(name=>['PropertyPane'+name+'Picker',factory])),'@apvee/spfx-react-toolkit/property-pane/file-folder-picker':{PropertyPaneFileFolderPicker:factory},SpFxReactToolkitTestWebPartStrings:{},'./components/WebPartShellDemoHost':{WebPartShellDemoHost:()=>undefined},'./components/SpFxReactToolkitTest':{default:()=>undefined}});
 return new (h.load(path.join(base,'SpFxReactToolkitTestWebPart.ts')).default)(); }
function fields(part){return part.getPropertyPaneConfiguration().pages[0].groups.flatMap(group=>group.groupFields)}
const selected={kind:'file',tenantId:'11111111-1111-4111-8111-111111111111',endpoint:'https://tenant.sharepoint.com/_api/v2.0',driveId:'b!d',itemId:'ONE',name:'Report.pdf'};
test('native sample exposes a real factory with independent hosted modes and supported media sizes',()=>{const part=loadPart();part.properties={fileFolderValues:[selected],fileFolderItemTypes:'all',fileFolderMultiselect:true,fileFolderSize:'large',pickerPresentation:'combobox'};
 const props=fields(part).find(field=>field.targetProperty==='fileFolderValues').properties;assert.equal(props.presentation,undefined);assert.equal(fields(part).some(f=>f.targetProperty==='fileFolderPresentation'),false);assert.equal(props.itemTypes,'all');assert.equal(props.multiselect,true);assert.equal(props.size,'large');assert.deepEqual(props.value,[selected]);assert.deepEqual(JSON.parse(JSON.stringify(props.value)),[selected]);assert.equal(props.scope.kind,'siteCollection');});
test('native sample source cascade clears dependent refs only; presentation and labels retain refs',()=>{const part=loadPart();part.properties={fileFolderValues:[selected],pickerSites:[],pickerLists:[]};part.onPropertyPaneFieldChanged('fileFolderSize','medium','large');assert.deepEqual(part.properties.fileFolderValues,[selected]);const revision=fields(part).find(f=>f.targetProperty==='fileFolderValues').properties.valueRevision;
 part.onPropertyPaneFieldChanged('fileFolderScope','siteCollection','library');assert.deepEqual(part.properties.fileFolderValues,[]);assert.ok(fields(part).find(f=>f.targetProperty==='fileFolderValues').properties.valueRevision>revision);});
test('missing dependent library disables native picker; selected library uses stable toolkit selector',()=>{const part=loadPart();part.properties={fileFolderScope:'library'};assert.equal(fields(part).find(f=>f.targetProperty==='fileFolderValues').properties.disabled,true);const site={kind:'site',url:'https://tenant.sharepoint.com/sites/hr'},list={kind:'list',webUrl:site.url,listId:'11111111-1111-4111-8111-111111111111'};part.properties={fileFolderScope:'library',pickerSites:[site],pickerLists:[list]};const props=fields(part).find(f=>f.targetProperty==='fileFolderValues').properties;assert.equal(props.disabled,false);assert.deepEqual(props.scope,{kind:'library',webUrl:site.url,list:{kind:'id',id:list.listId}});});
test('registry and live sample expose actual hosted control, consumer values and starting location',()=>{const panel=fs.readFileSync(path.join(base,'components/panels/FileFolderPickerPanel.tsx'),'utf8'),registry=fs.readFileSync(path.join(base,'components/demoRegistry.ts'),'utf8');for(const symbol of ['FileFolderPicker','PropertyPaneFileFolderPicker'])assert.match(registry,new RegExp("symbol: '"+symbol+"'"));assert.match(panel,/<FileFolderPicker /);for(const prop of ['scope','initialLocation','itemTypes','fileCategory','multiselect','size','selectionLimit','extensions','value','initialOptions'])assert.match(panel,new RegExp(prop+'=\\{'));assert.doesNotMatch(panel,/localStorage|sessionStorage|\/me\/drive|createIfMissing/);});

test('native OneDrive source uses current user without URL entry',()=>{const part=loadPart();part.properties={fileFolderScope:'oneDrive'};const configured=fields(part);assert.deepEqual(configured.find(f=>f.targetProperty==='fileFolderValues').properties.scope,{kind:'oneDrive',folder:undefined});assert.equal(configured.some(f=>f.targetProperty==='fileFolderBaseUrl'),false);});

test('new native configuration uses focused without legacy SDK options', () => {
 const part=loadPart(), configured=fields(part), props=configured.find(f=>f.targetProperty==='fileFolderValues').properties;
 assert.equal(props.preset,'focused');
 for(const name of ['graphEnabled','initialView','navigationPreset','visibleLocations','localFilesMode','title','layout']) assert.equal(props[name],undefined,name);
 for(const key of ['fileFolderGraph','fileFolderMyFiles','fileFolderNavigationVisible','fileFolderLayout','fileFolderDeviceUpload']) assert.equal(configured.some(f=>f.targetProperty===key),false,key);
 assert.deepEqual(configured.find(f=>f.targetProperty==='fileFolderPreset').properties.options.map(o=>o.key),['focused','extended']);
});
test('native preset menu offers only focused and extended for saved legacy configurations without losing values', () => {
 for (const preset of [undefined, 'legacy']) {
  const part=loadPart(); part.properties={fileFolderPreset:preset,fileFolderGraph:false,fileFolderTitle:'Saved title',fileFolderLayout:'tiles',fileFolderRecent:true,fileFolderUpload:true,fileFolderCreateFolder:true,fileFolderValues:[selected]};
  const picker=()=>fields(part).find(f=>f.targetProperty==='fileFolderValues').properties;
  const menu=()=>fields(part).find(f=>f.targetProperty==='fileFolderPreset').properties;
  const before=picker();assert.equal(before.preset,'focused');assert.equal(menu().selectedKey,'focused');
  assert.deepEqual(menu().options.map(o=>o.key),['focused','extended']);
  for(const name of ['graphEnabled','title','layout','visibleLocations'])assert.equal(before[name],undefined,name);
  part.properties.fileFolderPreset='extended';part.onPropertyPaneFieldChanged('fileFolderPreset',preset,'extended');
  const after=picker();assert.equal(after.preset,'extended');assert.equal(menu().selectedKey,'extended');
  assert.equal(after.uploadEnabled,true);assert.equal(after.createFolderEnabled,true);assert.deepEqual(after.value,[selected]);assert.equal(after.key,before.key);assert.equal(after.valueRevision,before.valueRevision);
  assert.equal(part.properties.fileFolderTitle,'Saved title');assert.equal(part.properties.fileFolderGraph,false);
 }
});
test('React and native new demos share preset intents and preset-only edits retain selection and models',async()=>{
 const {createPickerHarness}=require('./fixtures/picker-source-harness.cjs');
 const React=require('react'),{Simulate}=require('react-dom/test-utils');let observed,mounts=0,changes=0;
 const models=[{value:selected,text:'Report.pdf'}];
 const Probe=p=>{observed=p;React.useEffect(()=>{mounts++;},[]);return React.createElement('button',{onClick:()=>{changes++;p.onChange({reason:'pick',values:[selected],items:models,unresolvedValues:[]});}},'Seed selection');};
 const h=createPickerHarness({
 '@apvee/spfx-react-toolkit':{useSPFxPageContext:()=>({web:{absoluteUrl:'https://tenant.sharepoint.com/sites/hr'}}),FileFolderPicker:Probe},
 '../shared/DemoCard':{DemoCard:p=>React.createElement('div',undefined,React.createElement('section',{'data-config':true},p.controls),p.commands,React.createElement('section',{'data-diagnostics':true},p.diagnostics),p.children)},
 '../shared/DemoAccordionPanel':{DemoAccordionPanel:p=>React.createElement('div',undefined,p.children)},
 });
 const Panel=h.load(path.join(base,'components/panels/FileFolderPickerPanel.tsx')).default,m=h.mountComponent(()=>React.createElement(Panel));
 const input=label=>{const l=[...m.container.querySelectorAll('label')].find(l=>l.textContent===label);assert.ok(l,label);return m.container.querySelector(`[id="${l.htmlFor}"]`);};
 const change=(label,value)=>{const el=input(label);h.act(()=>{el.value=value;Simulate.change(el);});};
 try{
  const native=fields(loadPart()).find(f=>f.targetProperty==='fileFolderValues').properties;
  for(const key of ['preset','itemTypes','multiselect','uploadEnabled','createFolderEnabled','searchable','colorfulTagAvatars'])assert.equal(observed[key],native[key],key);
  const metadata=require('./fixtures/file-folder-picker/metadata.cjs').setup();
  const resolved=await metadata.load('scope').resolveFileFolderPickerScope({kind:'siteCollection'},undefined,metadata.context,metadata.cache);
  const compile=metadata.load('config').compileFileFolderPickerConfig,channel={origin:'https://tenant.sharepoint.com',channelId:'demo-parity'};
  assert.deepEqual(compile(resolved,observed,channel),compile(resolved,native,channel),'same library resolver yields identical hosted SDK JSON');
  assert.equal(observed.preset,'focused');h.act(()=>Simulate.click([...m.container.querySelectorAll('button')].find(b=>b.textContent==='Seed selection')));
  const savedModels=observed.initialOptions;
  const colorful = input('Colorful tag avatars');
  h.act(()=>{colorful.checked=true;Simulate.change(colorful);});
  assert.equal(observed.colorfulTagAvatars,true);assert.deepEqual(observed.value,[selected]);assert.equal(observed.initialOptions,savedModels);
  change('File extensions (comma separated)', '.pdf');assert.deepEqual(observed.extensions,['.pdf']);
  change('File category','images');assert.equal(observed.fileCategory,'images');assert.equal(observed.extensions,undefined);assert.equal(input('File extensions (comma separated)').value,'');assert.equal(input('File extensions (comma separated)').disabled,true);
  change('Selectable items','all');assert.equal(observed.fileCategory,'images');assert.equal(observed.itemTypes,'all');assert.deepEqual(observed.value,[selected]);assert.equal(observed.initialOptions,savedModels);
  change('Selectable items','folders');assert.equal(observed.fileCategory,undefined);assert.equal(observed.extensions,undefined);assert.deepEqual(observed.value,[selected]);assert.equal([...m.container.querySelectorAll('label')].some(l=>l.textContent==='File category'),false);
  change('Selectable items','files');assert.equal(observed.fileCategory,undefined);assert.equal(input('File extensions (comma separated)').disabled,false);
  change('File category','documents');change('File category','');assert.equal(observed.fileCategory,undefined);assert.deepEqual(observed.extensions,[]);
  change('Preset','extended');assert.equal(observed.preset,'extended');assert.deepEqual(observed.value,[selected]);assert.equal(observed.initialOptions,savedModels);assert.equal(mounts,1);assert.equal(changes,1);
  for(const key of ['graphEnabled','claimsChallengeEnabled','visibleLocations','initialView','navigationPreset','localFilesMode','title','layout'])assert.equal(observed[key],undefined,key);
  assert.doesNotMatch(m.container.textContent,/Optional Graph|Request Recent|Initial native layout|Native tray|On this device label/);
  assert.equal(m.container.querySelector('[data-diagnostics]').querySelector('input,select,textarea'),null,'diagnostics are read-only');
  change('Initial web URL (navigation only)','https://tenant.sharepoint.com/sites/hr');change('Initial library title (navigation only)','Documents');change('Initial folder (navigation only)','Policies');
  const apply=()=>h.act(()=>Simulate.click([...m.container.querySelectorAll('button')].find(b=>b.textContent==='Apply source and clear demo selection')));
  apply();const initial=observed.initialLocation;assert.equal(initial.folder,'Policies');change('Preset','focused');assert.equal(observed.initialLocation,initial);
  change('Source','oneDrive');assert.equal(input('Initial folder (navigation only)').value,'');assert.equal([...m.container.querySelectorAll('label')].some(l=>l.textContent==='Initial web URL (navigation only)'),false);
  apply();assert.deepEqual(observed.scope,{kind:'oneDrive',folder:undefined});assert.equal(observed.initialLocation,undefined);

 }finally{m.unmount();h.close();}
});

test('native options and diagnostics are compact and only relevant limits/filters are shown', () => {
 const part=loadPart();const groups=part.getPropertyPaneConfiguration().pages[0].groups;
 for(const name of ['File picker options','File picker diagnostics'])assert.equal(groups.find(g=>g.groupName===name).isCollapsed,true);
 assert.equal(fields(part).some(f=>f.targetProperty==='fileFolderSelectionLimit'),false);
 part.properties={fileFolderMultiselect:true,fileFolderItemTypes:'folders',fileFolderSelectionLimit:'3',fileFolderExtensions:'.pdf'};
 const configured=fields(part),props=configured.find(f=>f.targetProperty==='fileFolderValues').properties;
 assert.equal(configured.some(f=>f.targetProperty==='fileFolderSelectionLimit'),true);assert.equal(configured.some(f=>f.targetProperty==='fileFolderExtensions'),false);
 assert.equal(props.selectionLimit,3);assert.equal(props.extensions,undefined);
});
test('legacy read adapter remains available independently of the native preset menu', () => {
 const h=createHarness({});const {fileFolderDemoExperience}=h.load(path.join(base,'components/shared/fileFolderPickerNavigation.ts'));
 const settings={fileFolderInitialView:'recent',fileFolderMyFiles:true,fileFolderQuickAccess:true,fileFolderSite:true,fileFolderNavigationPreset:'oneDrive',fileFolderSiteUrl:'https://tenant-my.sharepoint.com/personal/user',fileFolderFromLink:'pivot',fileFolderDeviceUpload:true,fileFolderDeviceLabel:'Saved device label'};
 const picker=scope=>fileFolderDemoExperience(settings,scope);
 assert.deepEqual(picker({kind:'oneDrive'}).visibleLocations,['myFiles','site']);assert.equal(picker({kind:'oneDrive'}).initialView,'recent');assert.equal(picker({kind:'oneDrive'}).localFilesMode,'disabled');
 settings.fileFolderUpload=true;assert.equal(picker({kind:'oneDrive'}).localFilesMode,'upload');assert.equal(picker({kind:'oneDrive'}).localFilesLabel,'Saved device label');
 assert.equal(picker({kind:'oneDrive'}).navigationPreset,'oneDrive');assert.equal(picker({kind:'oneDrive'}).fromLinkMode,'pivot');
 assert.equal(picker({kind:'oneDrive',folder:'Projects'}).initialView,'location');
 assert.deepEqual(picker({kind:'siteCollection'}).visibleLocations,['quickAccess','site']);assert.equal(picker({kind:'siteCollection'}).initialView,'location');
});

test('native source changes clear optional starting location while preset changes retain it and values', () => {
 const part=loadPart();part.properties={fileFolderValues:[selected],fileFolderInitialWeb:'https://tenant.sharepoint.com/sites/hr',fileFolderInitialLibrary:'Documents',fileFolderInitialFolder:'Policies'};
 const picker=()=>fields(part).find(f=>f.targetProperty==='fileFolderValues').properties;
 const before=picker();part.properties.fileFolderPreset='extended';part.onPropertyPaneFieldChanged('fileFolderPreset','focused','extended');
 assert.deepEqual(picker().initialLocation,before.initialLocation);assert.deepEqual(picker().value,[selected]);assert.equal(picker().valueRevision,before.valueRevision);
 part.properties.fileFolderScope='oneDrive';part.onPropertyPaneFieldChanged('fileFolderScope','siteCollection','oneDrive');
 assert.equal(picker().initialLocation,undefined);assert.deepEqual(picker().value,[]);assert.ok(picker().valueRevision>before.valueRevision);
 for(const name of ['fileFolderInitialWeb','fileFolderInitialLibrary'])assert.equal(fields(part).some(f=>f.targetProperty===name),false);
 assert.equal(fields(part).some(f=>f.targetProperty==='fileFolderInitialFolder'),true);
});

test('native effective source changes clear stale starting location without altering reset rules', () => {
 const site={kind:'site',url:'https://tenant.sharepoint.com/sites/hr'};
 const list={kind:'list',webUrl:site.url,listId:'11111111-1111-4111-8111-111111111111'};
 const replacements={pickerSites:[{...site,url:'https://tenant.sharepoint.com/sites/finance'}],pickerLists:[{...list,listId:'22222222-2222-4222-8222-222222222222'}],fileFolderFolder:'New folder'};
 for(const scope of ['siteCollection','library','folder','oneDrive']) for(const property of ['pickerSites','pickerLists','fileFolderFolder']) {
  const part=loadPart();part.properties={fileFolderScope:scope,pickerSites:[site],pickerLists:[list],fileFolderFolder:'Original folder',fileFolderValues:[selected],fileFolderInitialWeb:site.url,fileFolderInitialLibrary:'Documents',fileFolderInitialFolder:'Original folder/Start'};
  const picker=()=>fields(part).find(f=>f.targetProperty==='fileFolderValues').properties;
  const before=picker(),oldValue=part.properties[property];part.properties[property]=replacements[property];part.onPropertyPaneFieldChanged(property,oldValue,replacements[property]);
  const after=picker(),dependent=property==='pickerSites' ? scope!=='oneDrive' : property==='pickerLists' ? scope==='library'||scope==='folder' : scope==='folder'||scope==='oneDrive';
  assert.equal(after.initialLocation===undefined,dependent,`${scope} / ${property}`);
  if(!dependent)assert.deepEqual(after.initialLocation,before.initialLocation);
  assert.deepEqual(after.value,[],'existing selection reset rules remain intact');assert.equal(after.valueRevision,before.valueRevision+1);assert.equal(after.key,before.key);
 }
});

test('native cosmetic and same-identity source updates preserve optional starting location and selection', () => {
 const site={kind:'site',url:'https://tenant.sharepoint.com/sites/hr'};
 const list={kind:'list',webUrl:site.url,listId:'11111111-1111-4111-8111-111111111111'};
 for(const [property,oldValue,newValue] of [
  ['pickerSites',[site],[{...site,url:site.url.toUpperCase()+'/',text:'Renamed site'}]],
  ['pickerLists',[list],[{...list,listId:list.listId.toUpperCase(),text:'Renamed library'}]],
  ['fileFolderFolder','Original folder','Original folder'],
  ['fileFolderPreset','focused','extended'],
 ]) {
  const part=loadPart();part.properties={fileFolderScope:'folder',pickerSites:[site],pickerLists:[list],fileFolderFolder:'Original folder',fileFolderValues:[selected],fileFolderInitialWeb:site.url,fileFolderInitialLibrary:'Documents',fileFolderInitialFolder:'Original folder/Start'};
  const picker=()=>fields(part).find(f=>f.targetProperty==='fileFolderValues').properties;
  const before=picker();part.properties[property]=newValue;part.onPropertyPaneFieldChanged(property,oldValue,newValue);
  const after=picker();assert.deepEqual(after.initialLocation,before.initialLocation,property);assert.deepEqual(after.value,before.value);assert.equal(after.valueRevision,before.valueRevision);assert.equal(after.key,before.key);
 }
});


test('native collapsed groups remain reachable and every source has nonempty groups without mutating saved values', () => {
 for (const scope of ['siteCollection', 'library', 'folder', 'oneDrive']) {
  const part = loadPart();
  part.properties = { fileFolderScope: scope, fileFolderValues: [selected], fileFolderItemTypes: 'folders', fileFolderMultiselect: false };
  const saved = JSON.stringify(part.properties);
  const page = part.getPropertyPaneConfiguration().pages[0];
  assert.equal(page.displayGroupsAsAccordion, true, 'collapsed groups need an interactive accordion header');
  for (const group of page.groups) assert.ok(group.groupFields.length > 0, `${scope}: no empty group heading`);
  const options = page.groups.find(group => group.groupName === 'File picker options');
  assert.ok(options.groupFields.some(field => field.targetProperty === 'fileFolderHint'));
  assert.ok(options.groupFields.some(field => field.targetProperty === 'fileFolderSearchable'));
  assert.equal(JSON.stringify(part.properties), saved, 'building the pane must retain saved configuration and selection');
 }
});

test('native category menu clears incompatible extensions while preserving references and revision', () => {
 const part=loadPart();part.properties={fileFolderValues:[selected],fileFolderExtensions:'.pdf'};
 const picker=()=>fields(part).find(f=>f.targetProperty==='fileFolderValues').properties;
 const before=picker();
 assert.deepEqual(fields(part).find(f=>f.targetProperty==='fileFolderCategory').properties.options.map(o=>o.key),['','images','video','audio','media','documents']);
 for(const category of ['images','video','audio','media','documents']) {
  part.properties.fileFolderCategory=category;part.onPropertyPaneFieldChanged('fileFolderCategory',undefined,category);
  assert.equal(part.properties.fileFolderExtensions,undefined);assert.equal(picker().fileCategory,category);assert.equal(picker().extensions,undefined);
  assert.equal(fields(part).find(f=>f.targetProperty==='fileFolderExtensions').properties.disabled,true);
  assert.deepEqual(picker().value,[selected]);assert.equal(picker().valueRevision,before.valueRevision);assert.equal(picker().key,before.key);
 }
 part.properties.fileFolderItemTypes='all';part.onPropertyPaneFieldChanged('fileFolderItemTypes','files','all');
 assert.equal(picker().fileCategory,'documents');assert.equal(picker().itemTypes,'all');
 part.properties.fileFolderItemTypes='folders';part.onPropertyPaneFieldChanged('fileFolderItemTypes','all','folders');
 assert.equal(part.properties.fileFolderCategory,undefined);assert.equal(picker().fileCategory,undefined);assert.equal(fields(part).some(f=>f.targetProperty==='fileFolderCategory'),false);
 assert.deepEqual(picker().value,[selected]);assert.equal(picker().valueRevision,before.valueRevision);
 part.properties.fileFolderItemTypes='files';part.properties.fileFolderCategory='';part.onPropertyPaneFieldChanged('fileFolderCategory','documents','');
 assert.equal(picker().fileCategory,undefined);assert.equal(fields(part).find(f=>f.targetProperty==='fileFolderExtensions').properties.disabled,false);
});
