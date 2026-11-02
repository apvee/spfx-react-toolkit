const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createContextFixture } = require('./fixtures/executable-examples/context-scenarios.cjs');
const hubTitle = 'Start a scope bound hub association lookup';
const hubOwner = 'packages/spfx-react-toolkit/src/hooks/useSPFxHubSiteInfo.ts#useSPFxHubSiteInfo';
const snapshots = [
  ['Context','Inspect the current native host identity',/WebPart.*context-one/],
  ['PageContext','Inspect selected native page metadata',/Project intranet.*Giulia Rossi/],
  ['InstanceInfo','Inspect the provider instance identity',/context-one.*WebPart/],
  ['DisplayMode','Observe the native Read and Edit mode',/ModeRead.*Readtrue.*Editfalse/],
  ['IsEdit','Observe the native edit mode shortcut',/ModeRead.*Is editfalse/],
  ['EnvironmentInfo','Observe the acquired application host',/HostSharePoint.*Teamsfalse.*Officefalse/],
  ['UserInfo','Inspect the current user display metadata',/Giulia Rossi.*giulia@contoso.com.*External userfalse/],
  ['SiteInfo','Inspect the current web and site collection',/Project intranet.*ClassificationInternal/],
  ['ListInfo','Inspect the optional current list context',/No list context/],
  ['LocaleInfo','Compare host cultures and seasonal site times',/it-IT.*en-US.*Europe\/Berlin/],
  ['Teams','Observe application SDK readiness and theme',/SDK context acquiredfalse.*SDK themeUnavailable/],
  ['PageType','Inspect the supplied page classification',/Page typesitePage.*Moderntrue/],
  ['CorrelationInfo','Inspect current diagnostic identifiers',/Correlation IDUnavailable.*Tenant IDUnavailable/],
  ['Permissions','Inspect native current page permission masks',/Manage Webfalse.*Web mask availablefalse/],
  ['ContainerInfo','Observe the native container and tracked size',/Host elementDIV.*WidthUnmeasured/],
  ['ContainerSize','Observe the measured container width category',/Width0.*Categorysmall.*Smalltrue/],
  ['ThemeInfo','Observe the native SPFx theme snapshot',/Native themeUnavailable/],
  ['Fluent9ThemeInfo','Observe the projected Fluent theme tokens',/Legacy isTeamsfalse.*Neutral background#ffffff/],
  ['ServiceScope','Inspect the native registered service scope',/Native service scopeAvailable/],
];
function owner(suffix) { return `packages/spfx-react-toolkit/src/${suffix === 'Context' ? 'core/context.internal.tsx' : `hooks/useSPFx${suffix === 'IsEdit' ? 'DisplayMode' : suffix}.ts`}#useSPFx${suffix}`; }
for (const [suffix, title, expected] of snapshots) test(`actual ${suffix} source case renders finite native metadata without business I/O`, async t => {
  const f = createContextFixture(t, title, owner(suffix)); await f.flush();
  assert.match(f.container.textContent, expected);
  assert.equal(f.calls.length, 0); assert.equal(f.entry.metadata.kind, 'snapshot');
  assert.equal(f.entry.metadata.host, 'webPart'); assert.equal(f.entry.metadata.contentPlacement, 'configuration');
  const subject = f.container.querySelector('[data-demo-result]'); assert.ok(subject);
  assert.equal(f.container.querySelectorAll('[data-demo-section]').length, 1);
  const shell = f.container.querySelector('[data-web-part-shell-surface]'); f.reset();
  assert.equal(f.container.querySelector('[data-web-part-shell-surface]'), shell, 'reset retains real outer Shell');
  assert.equal(f.calls.length, 0); assert.doesNotMatch(f.container.textContent, /OLD-HUB|accessToken|_private/);
});

test('Locale selects exactly one real seasonal time; reset and unavailable numeric mapping retain authoritative host metadata', async t => {
  const f = createContextFixture(t, snapshots[9][1], owner('LocaleInfo'));
  const shell = f.container.querySelector('[data-web-part-shell-surface]');
  assert.equal(f.container.querySelectorAll('time').length, 1); assert.equal(f.container.querySelector('time').dateTime, '2026-01-15T12:00:00.000Z');
  f.change('Period','summer'); assert.equal(f.container.querySelectorAll('time').length, 1); assert.equal(f.container.querySelector('time').dateTime, '2026-07-15T12:00:00.000Z');
  f.reset(); assert.equal(f.container.querySelector('time').dateTime, '2026-01-15T12:00:00.000Z'); assert.equal(f.container.querySelector('[data-web-part-shell-surface]'), shell);
  f.instance.context.pageContext.web.timeZoneInfo.id = 7; f.render();
  assert.equal(f.container.querySelectorAll('time').length, 0); assert.match(f.container.textContent, /IANA mapping unavailable/);
  f.instance.context.pageContext.web.timeZoneInfo.id = 13; f.render(); assert.match(f.container.textContent, /America\/Los_Angeles/); assert.equal(f.container.querySelectorAll('time').length, 1);
});

test('native list absence/fallback and finite diagnostic IDs update through actual hooks', async t => {
  const f = createContextFixture(t, snapshots[8][1], owner('ListInfo'));
  f.instance.context.pageContext.list = { id: 'cccccccc-0000-0000-0000-000000000001' }; f.render(); assert.match(f.container.textContent, /Unknown List.*List URLUnavailable.*TemplateUnavailable.*Document libraryfalse/);
  f.instance.context.pageContext.list = { id: 'cccccccc-0000-0000-0000-000000000001', title: 'Documents', serverRelativeUrl: '/sites/current/Docs', baseTemplate: 101 }; f.render(); assert.match(f.container.textContent, /Documents.*Document librarytrue/);
  f.select(snapshots[12][1], owner('CorrelationInfo')); f.instance.context.pageContext.site.correlationId = 'correlation-one'; f.instance.context.pageContext.aadInfo = { tenantId: 'tenant-one' }; f.render(); assert.match(f.container.textContent, /correlation-one.*tenant-one/);
});

test('native masks expose independent site/web/list checks without REST', async t => {
  const { Permission } = require('./fixtures/executable-examples/context-scenarios.cjs');
  const f = createContextFixture(t, snapshots[13][1], owner('Permissions'));
  const page = f.instance.context.pageContext;
  page.web.permissions = Permission.fullMask; page.site.permissions = Permission.viewListItems; page.list = { id: 'cccccccc-0000-0000-0000-000000000001', permissions: Permission.emptyMask }; f.render();
  assert.match(f.container.textContent, /Manage Webtrue.*Manage Liststrue.*Add List Itemstrue.*Site view list itemstrue.*List view list itemsfalse/);
  assert.equal(f.calls.length, 0);
});

test('actual Read/Edit and native SDK acquisition/theme events keep host classification separate', async t => {
  const f = createContextFixture(t, snapshots[3][1], owner('DisplayMode'));
  f.instance.displayMode = 2; f.render(); assert.match(f.container.textContent, /ModeEdit.*Readfalse.*Edittrue/);
  f.select(snapshots[4][1], owner('IsEdit')); assert.match(f.container.textContent, /ModeEdit.*Is edittrue/);
  f.select(snapshots[5][1], owner('EnvironmentInfo'));
  let resolve, theme;
  f.instance.context.sdks.microsoftTeams = { teamsJs: { app: { getContext: () => new Promise(done => { resolve = done; }), registerOnThemeChangeHandler: callback => { theme = callback; } } } }; f.render();
  assert.match(f.container.textContent, /HostSharePoint/);
  await f.flush(() => resolve({ app: { host: { name: 'Office' }, theme: 'default' } })); assert.match(f.container.textContent, /HostOffice.*Teamsfalse.*Officetrue/);
  f.select(snapshots[10][1], owner('Teams')); assert.match(f.container.textContent, /SDK context acquiredtrue/);
  await f.flush(() => theme('dark')); assert.match(f.container.textContent, /SDK themedark/);
  f.select(snapshots[17][1], owner('Fluent9ThemeInfo')); assert.match(f.container.textContent, /SDK themedark.*Legacy isTeamstrue/);
  f.instance.context.pageContext.web.absoluteUrl = 'http://localhost/workbench.aspx'; f.select(snapshots[5][1], owner('EnvironmentInfo')); assert.match(f.container.textContent, /HostLocal.*Localtrue.*Workbenchtrue/);
});

test('actual native theme subscription emits selected tokens and releases subscribers', async t => {
  const f = createContextFixture(t, snapshots[16][1], owner('ThemeInfo'));
  const theme = require('@fluentui/theme').createTheme({ palette: { themePrimary: '#ab1234' } });
  assert.equal(f.handlers.size, 1); f.h.act(() => { for (const fn of f.handlers) fn({ theme }); });
  assert.match(f.container.textContent, /Native themeAvailable.*#ab1234/);
  f.h.act(() => { for (const fn of f.handlers) fn({ theme: undefined }); }); assert.match(f.container.textContent, /Native themeUnavailable/);
  f.unmount(); assert.equal(f.handlers.size, 0);
});

test('container subjects use owner-document native observation and disconnect on departure', async t => {
  const f = createContextFixture(t, snapshots[14][1], owner('ContainerInfo'));
  const observers = [];
  const win = f.instance.domElement.ownerDocument.defaultView;
  const previous = Object.getOwnPropertyDescriptor(win,'ResizeObserver');
  win.ResizeObserver = class { constructor(callback) { this.callback=callback; observers.push(this); } observe(element) { this.element=element; } disconnect() { this.disconnected=true; } };
  t.after(() => { if (previous) Object.defineProperty(win,'ResizeObserver',previous); else delete win.ResizeObserver; });
  const next = f.host(); f.replace(next); await f.flush();
  assert.ok(observers.length); assert.equal(observers.at(-1).element, next.domElement);
  f.h.act(() => observers.at(-1).callback([{ contentRect: { width: 800, height: 400 } }])); assert.match(f.container.textContent, /Width800.*Height400/);
  f.select(snapshots[15][1], owner('ContainerSize')); await f.flush();
  f.h.act(() => observers.at(-1).callback([{ contentRect: { width: 480, height: 320 } }])); assert.match(f.container.textContent, /Width480.*Categorymedium.*Smallfalse.*Mediumtrue/);
  f.unmount(); assert.ok(observers.every(observer => observer.disconnected));
});

const legacyTitle = 'Check staged legacy site web and list permissions';
const scopedTitle = 'Check staged web list item and folder permissions';
const permissionOwner = owner('CrossSitePermissions');
async function settleMasks(f, low = 1, start = 0, mismatch = false) {
  const listId = 'cccccccc-0000-0000-0000-000000000001';
  for (let index = start; index < f.calls.length; index++) {
    const call = f.calls[index], mask = { High: '0', Low: String(low) };
    await f.flush(() => call.resolve(call.url.includes('GetFolder') ? { EffectiveBasePermissions: mask, ParentList: { Id: mismatch ? 'dddddddd-0000-0000-0000-000000000002' : listId } } : call.url.includes('?$select=Id') ? { Id: listId } : mask));
  }
}
test('legacy overload stages independent site/web/list inputs and repeats actual three-mask requests', async t => {
  const f = createContextFixture(t, legacyTitle, permissionOwner);
  f.change('Site URL','https://tenant.test/sites/target'); f.change('Selected web URL','https://tenant.test/sites/target/subweb'); f.change('List GUID','cccccccc-0000-0000-0000-000000000001');
  assert.equal(f.calls.length, 0); assert.match(f.container.textContent, /Idle/);
  f.click('Check / repeat'); await f.flush(); assert.equal(f.calls.length, 3); assert.match(f.container.textContent, /Loading legacy permissions/);
  assert.equal(f.calls[0].url,'https://tenant.test/sites/target/_api/site/rootweb/effectivebasepermissions');
  assert.match(f.calls[1].url,/\/subweb\/_api\/web\/effectivebasepermissions$/);
  await settleMasks(f, 0x40000803); assert.match(f.container.textContent,/Manage Webtrue.*Manage Liststrue.*Add List Itemstrue.*Site view list itemstrue.*List view list itemstrue/);
  const applied = f.container.querySelector('[aria-label="Legacy permission result"]').textContent;
  f.change('Site URL','https://tenant.test/sites/draft'); assert.equal(f.calls.length, 3); assert.equal(f.container.querySelector('[aria-label="Legacy permission result"]').textContent, applied);
  f.render(); assert.equal(f.calls.length, 3);
  f.click('Check / repeat'); await f.flush(); assert.equal(f.calls.length, 6); await settleMasks(f,0,3); assert.match(f.container.textContent,/Manage Webfalse.*Manage Listsfalse.*Add List Itemsfalse/);
  const input = f.container.querySelector('[aria-label="Site URL"]'); f.reset(); assert.equal(f.container.querySelector('[aria-label="Site URL"]'),input); assert.equal(input.value,''); assert.equal(f.calls.length,6); assert.match(f.container.textContent,/Idle/);
});

for (const value of ['', '   ', 'not a URL']) test(`legacy staged URL validation: ${JSON.stringify(value)}`, async t => {
  const f = createContextFixture(t, legacyTitle, permissionOwner); f.change('Site URL',value); f.click('Check / repeat'); await f.flush(); assert.equal(f.calls.length,0);
  assert.match(f.container.textContent, value.trim() ? /Error:/ : /Idle/); assert.equal(f.container.querySelectorAll('dd').length,0);
});
test('legacy invalid list GUID, transport failure and repeat recovery never publish grants on failure', async t => {
  const f=createContextFixture(t,legacyTitle,permissionOwner); f.change('Site URL','https://tenant.test/sites/target'); f.change('List GUID','invalid'); f.click('Check / repeat'); await f.flush(); assert.equal(f.calls.length,0); assert.match(f.container.textContent,/Error:/);
  f.change('List GUID',''); f.click('Check / repeat'); await f.flush(); await f.flush(() => f.calls[0].reject(new Error('Native transport failure'))); assert.match(f.container.textContent,/Error: Native transport failure/); assert.equal(f.container.querySelectorAll('dd').length,0);
  const start=f.calls.length; f.click('Check / repeat'); await f.flush(); await settleMasks(f,1,start); assert.doesNotMatch(f.container.textContent,/Error:/); assert.match(f.container.textContent,/Site view list itemstrue/);
});

for (const scope of ['web','list','item','folder']) for (const selector of (scope==='web'?['title']:['title','id','path'])) test(`scoped actual ${scope}/${selector} target, held loading, denied and repeat recovery`, async t => {
  const f=createContextFixture(t,scopedTitle,permissionOwner);
  f.change('Scope',scope); if(scope!=='web'){ f.change('List selector',selector); f.change('List value',selector==='id'?'cccccccc-0000-0000-0000-000000000001':selector==='path'?'Shared Documents':'Tasks'); }
  if(scope==='folder') f.change('Folder path','/sites/current/Shared Documents/Reports');
  assert.equal(f.calls.length,0); assert.equal(f.container.querySelector('[aria-label="Applied target"]'),null);
  f.click('Run / repeat'); await f.flush(); assert.match(f.container.textContent,/Loading scoped permissions/);
  const applied=JSON.parse(f.container.querySelector('[aria-label="Applied target"]').textContent); assert.equal(applied.scope,scope); assert.deepEqual(Object.keys(applied).sort(),scope==='web'?['scope']:['scope',selector==='id'?'listId':selector==='path'?'listPath':'listTitle',...(scope==='item'?['itemId']:scope==='folder'?['folderPath']:[])].sort());
  await settleMasks(f,0); assert.match(f.container.textContent,/Denied: view list items/);
  const count=f.calls.length, prior=f.container.querySelector('[aria-label="Applied target"]').textContent;
  f.change('Web URL','https://tenant.test/sites/target'); assert.equal(f.calls.length,count); assert.equal(f.container.querySelector('[aria-label="Applied target"]').textContent,prior); f.render(); assert.equal(f.calls.length,count);
  f.click('Run / repeat'); await f.flush(); await settleMasks(f,1,count); assert.match(f.container.textContent,/Allowed: view list items/); assert.ok(f.calls.slice(count).every(call=>call.url.startsWith('https://tenant.test/sites/target/')));
});

test('scoped native web mask completes without REST; invalid targets and folder/list mismatch are errors', async t => {
  const {Permission}=require('./fixtures/executable-examples/context-scenarios.cjs'); const f=createContextFixture(t,scopedTitle,permissionOwner);
  f.instance.context.pageContext.web.permissions=Permission.viewListItems; f.render(); f.click('Run / repeat'); await f.flush(); assert.equal(f.calls.length,0); assert.match(f.container.textContent,/Allowed: view list items/);
  f.change('Scope','item'); f.change('Item ID','0'); f.click('Run / repeat'); await f.flush(); assert.equal(f.calls.length,0); assert.match(f.container.textContent,/Error:/); assert.doesNotMatch(f.container.textContent,/Denied:|Allowed:/);
  f.change('Scope','list'); f.change('List selector','id'); f.change('List value','bad'); f.click('Run / repeat'); await f.flush(); assert.equal(f.calls.length,0); assert.match(f.container.textContent,/Error:/);
  f.change('Scope','folder'); f.change('List selector','title'); f.change('List value','Docs'); f.change('Folder path','/sites/current/Docs/Reports'); f.click('Run / repeat'); await f.flush(); await settleMasks(f,1,0,true); assert.match(f.container.textContent,/Error:/); assert.doesNotMatch(f.container.textContent,/Denied:|Allowed:/);
  const start=f.calls.length; f.click('Run / repeat'); await f.flush(); await settleMasks(f,1,start); assert.match(f.container.textContent,/Allowed:/);
});

for (const title of [legacyTitle,scopedTitle]) for (const action of ['reset','departure','user','host','repeat']) for (const reject of [false,true]) test(`${title}: ${action} suppresses held ${reject?'error':'grant'}`, async t => {
  const f=createContextFixture(t,title,permissionOwner,{directProvider:true}); const legacy=title===legacyTitle, button=legacy?'Check / repeat':'Run / repeat';
  if(legacy) f.change('Site URL','https://tenant.test/sites/target');
  f.click(button); await f.flush(); const old=f.calls.slice(), first=f.observations.length;
  if(action==='reset') f.reset(); else if(action==='departure') f.unmount(); else if(action==='user'){f.instance.context.pageContext.user.loginName='replacement';f.render();} else if(action==='host') f.replace(f.host('context-replacement')); else { f.change(legacy?'Site URL':'Web URL','https://tenant.test/sites/replacement'); f.click(button); }
  await f.flush(() => old.forEach(call=>reject?call.reject(new Error('OLD-ERROR')):call.resolve({High:'0',Low:'1073743875'})));
  for(const record of f.observations.slice(first)) assert.doesNotMatch(record.text,/OLD-ERROR|Allowed:|Manage Webtrue/);
});

test('two actual providers in one source/cache universe keep initiated Hub results isolated', async t => {
  const f=createContextFixture(t,hubTitle,hubOwner,{directProvider:true}); const peer=f.mountPeer();
  f.click('Start / repeat lookup'); peer.click('Start / repeat lookup'); await f.flush(); assert.equal(f.calls.length,2);
  await f.flush(()=>f.calls[0].resolve({value:JSON.stringify({url:'https://tenant.test/ONLY-FIRST'})})); assert.match(f.container.textContent,/ONLY-FIRST/); assert.doesNotMatch(peer.container.textContent,/ONLY-FIRST/);
  await f.flush(()=>f.calls[1].resolve({value:JSON.stringify({url:'https://tenant.test/ONLY-PEER'})})); assert.match(peer.container.textContent,/ONLY-PEER/); assert.doesNotMatch(f.container.textContent,/ONLY-PEER/);
  peer.unmount();
});

for (const title of [legacyTitle,scopedTitle]) test(`${title}: fresh child attempt may share only an identical pending native request; settled repeat is fresh`, async t => {
  const f=createContextFixture(t,title,permissionOwner,{directProvider:true}), legacy=title===legacyTitle;
  if(legacy) f.change('Site URL','https://tenant.test/sites/target');
  const button=legacy?'Check / repeat':'Run / repeat', label=legacy?'Legacy permission result':'Scoped permission result';
  f.click(button); await f.flush(); const calls=f.calls.length, original=f.container.querySelector(`[aria-label="${label}"]`);
  f.click(button); await f.flush();
  assert.notEqual(f.container.querySelector(`[aria-label="${label}"]`),original,'explicit command mounts a fresh query child');
  assert.equal(f.calls.length,calls,'identical inflight promise is shared only in matching native scope');
  await settleMasks(f,1); assert.match(f.container.textContent,legacy?/Site view list itemstrue/:/Allowed:/);
  f.click(button); await f.flush(); assert.equal(f.calls.length,calls+(legacy?2:1),'settled repeat requests masks again');
  await settleMasks(f,0,calls); assert.match(f.container.textContent,legacy?/Site view list itemsfalse/:/Denied:/);
});

test('two permission providers in the same actual module/cache do not share native promises or decisions', async t => {
  const f=createContextFixture(t,scopedTitle,permissionOwner,{directProvider:true}), peer=f.mountPeer();
  f.click('Run / repeat'); peer.click('Run / repeat'); await f.flush(); assert.equal(f.calls.length,2);
  await f.flush(()=>f.calls[0].resolve({High:'0',Low:'1'})); assert.match(f.container.textContent,/Allowed:/); assert.doesNotMatch(peer.container.textContent,/Allowed:/);
  await f.flush(()=>f.calls[1].resolve({High:'0',Low:'0'})); assert.match(peer.container.textContent,/Denied:/); assert.match(f.container.textContent,/Allowed:/); peer.unmount();
});

test('integrated native boundary holds each request independently, captures its mode and rejects unknown services', async t => {
  const f=createContextFixture(t,snapshots[0][1],owner('Context'));
  const directory=path.resolve(__dirname,'fixtures/demo-presentation');
  const sdk=f.h.load(path.join(directory,'spfx-services.ts'));
  const native=f.h.load(path.join(directory,'trim-transports.ts'));
  const client=native.nativeFixtureService(sdk.SPHttpClient.serviceKey,{});
  assert.throws(()=>native.nativeFixtureService({},{}),/Unknown native/);
  window.trimFixture.configure({mode:'allowed',hold:true});
  const first=client.get('https://tenant.test/_api/web/effectivebasepermissions');
  window.trimFixture.configure({mode:'denied'});
  native.contextRequests[0].resolve();
  assert.deepEqual(await (await first).json(),{High:'0',Low:'31'},'held read preserves the initiation mode');
  const second=client.get('https://tenant.test/_api/web/hubsitedata(false)');
  const third=client.get('https://tenant.test/_api/web/hubsitedata(false)');
  const rejection=assert.rejects(third,/obsolete native transport/);
  native.contextRequests[2].reject(new Error('obsolete native transport'));
  native.contextRequests[1].resolve({value:JSON.stringify({url:'https://tenant.test/INDEPENDENT'})});
  assert.deepEqual(await (await second).json(),{value:JSON.stringify({url:'https://tenant.test/INDEPENDENT'})});
  await rejection;
});

for (const title of [hubTitle,legacyTitle,scopedTitle]) test(`${title}: Reset clears retired scope feedback and restores the local idle baseline`, async t => {
  const f=createContextFixture(t,title,title===hubTitle?hubOwner:permissionOwner,{directProvider:true});
  if(title===legacyTitle) f.change('Site URL','https://tenant.test/sites/target');
  f.click(title===hubTitle?'Start / repeat lookup':title===legacyTitle?'Check / repeat':'Run / repeat'); await f.flush();
  f.instance.context.pageContext.user.loginName='replacement'; f.render();
  assert.match(f.container.textContent,/Scope changed/);
  const reads=f.calls.length; f.reset();
  assert.match(f.container.textContent,/Idle\./); assert.doesNotMatch(f.container.querySelector('[data-demo-result]').textContent,/Scope changed/);
  assert.equal(f.calls.length,reads,'reset does not issue business I/O');
});
test('actual emitted Hub child hides old results on its own commit after native in-place mutation', async t => {
  const f = createContextFixture(t, hubTitle, hubOwner, { directProvider: true, parentOnly: process.env.CONTEXT_PARENT_ONLY_MUTATION === '1' });
  assert.equal(f.calls.length, 0);
  f.click('Start / repeat lookup'); await f.flush(); assert.equal(f.calls.length, 1);
  const renders = f.parentRenders, first = f.observations.length;
  f.instance.context.pageContext.legacyPageContext.hubSiteId = 'bbbbbbbb-0000-0000-0000-000000000002';
  await f.flush(() => f.calls[0].resolve({ value: JSON.stringify({ url: 'https://tenant.test/OLD-HUB' }) }));
  assert.ok(f.observations.slice(first).some(record => record.parentRenders === renders), 'held completion commits without parent render');
  for (const record of f.observations.slice(first)) assert.doesNotMatch(record.text, /OLD-HUB/);
  assert.match(f.container.textContent, /Scope changed.*Start/);
  assert.equal(f.calls.length, 2, 'native refresh scheduled after Start before retirement is preserved');
  t.diagnostic('Hub ID mutation without parent render: initial GET=1; old completion plus native refresh GETs=2; obsolete URL commits=0.');
  await f.flush(() => f.calls[1].reject(new Error('OLD-REFRESH')));
  assert.doesNotMatch(f.container.textContent, /OLD-REFRESH/);
});

for (const failure of [false, true]) for (const mutation of ['hub', 'current-hub', 'legacy-url', 'web', 'user']) {
  test(`Hub child scope isolation without parent render: ${mutation}, ${failure ? 'reject' : 'resolve'}`, async t => {
    const f = createContextFixture(t, hubTitle, hubOwner, { directProvider: true });
    f.click('Start / repeat lookup'); await f.flush();
    const page = f.instance.context.pageContext;
    const start = f.observations.length;
    if (mutation === 'hub') page.legacyPageContext.hubSiteId = 'bbbbbbbb-0000-0000-0000-000000000002';
    if (mutation === 'current-hub') page.legacyPageContext.isHubSite = true;
    if (mutation === 'legacy-url') page.webAbsoluteUrl = 'https://tenant.test/sites/B';
    if (mutation === 'web') page.web.absoluteUrl = 'https://tenant.test/sites/B';
    if (mutation === 'user') page.user.loginName = 'another-user';
    await f.flush(() => failure ? f.calls[0].reject(new Error('OLD-ERROR')) : f.calls[0].resolve({ value: JSON.stringify({ url: 'https://tenant.test/OLD-HUB' }) }));
    for (const record of f.observations.slice(start)) assert.doesNotMatch(record.text, /OLD-HUB|OLD-ERROR/);
    assert.match(f.container.textContent, /Scope changed/);
    t.diagnostic(`Hub child ${mutation}/${failure ? 'reject' : 'resolve'}: native GETs after old settlement=${f.calls.length}; obsolete display commits=0.`);
    // Observed retirement remains retired when the original scalar returns.
    page.legacyPageContext.hubSiteId = 'aaaaaaaa-0000-0000-0000-000000000001';
    page.legacyPageContext.isHubSite = false; delete page.webAbsoluteUrl;
    page.web.absoluteUrl = 'https://tenant.test/sites/current'; page.user.loginName = 'giulia'; f.render();
    assert.match(f.container.textContent, /Scope changed/);
    f.click('Start / repeat lookup'); await f.flush();
    const fresh = f.calls.at(-1);
    await f.flush(() => fresh.resolve({ value: JSON.stringify({ url: 'https://tenant.test/NEW-HUB' }) }));
    assert.match(f.container.textContent, /NEW-HUB/); assert.doesNotMatch(f.container.textContent, /OLD-/);
    for (const call of f.calls.slice(1,-1)) await f.flush(() => call.reject(new Error('OLD-REFRESH')));
    assert.match(f.container.textContent, /NEW-HUB/); assert.doesNotMatch(f.container.textContent, /OLD-/);
  });
}

for (const action of ['reset', 'departure', 'context', 'page', 'scope', 'client', 'new-start']) for (const failure of [false, true]) {
  test(`Hub ${action} retires old ${failure ? 'rejection' : 'URL'} before and after new Start`, async t => {
    const f = createContextFixture(t, hubTitle, hubOwner, { directProvider: true });
    f.click('Start / repeat lookup'); await f.flush(); const old = f.calls[0];
    if (action === 'reset') f.reset();
    else if (action === 'departure') f.unmount();
    else if (action === 'context') f.replace(f.host('context-two'));
    else if (action === 'page') f.replace(f.host('context-one', f.page()));
    else if (action === 'scope' || action === 'client') {
      const next = f.host(); f.instance.context.serviceScope = next.context.serviceScope; f.instance.context.pageContext = next.context.pageContext; f.render();
    } else f.click('Start / repeat lookup');
    const first = f.observations.length;
    await f.flush(() => failure ? old.reject(new Error('OLD-ERROR')) : old.resolve({ value: JSON.stringify({ url: 'https://tenant.test/OLD-HUB' }) }));
    for (const record of f.observations.slice(first)) assert.doesNotMatch(record.text, /OLD-HUB|OLD-ERROR/);
    if (action !== 'departure') {
      if (action !== 'new-start') f.click('Start / repeat lookup');
      await f.flush(); await f.flush(() => f.calls.at(-1).resolve({ value: JSON.stringify({ url: 'https://tenant.test/NEW-HUB' }) }));
      assert.match(f.container.textContent, /NEW-HUB/);
    }
  });
}

test('Hub finite states, repeat, no-association/current-hub shortcuts and provider isolation', async t => {
  const a = createContextFixture(t, hubTitle, hubOwner);
  const b = createContextFixture(t, hubTitle, hubOwner);
  a.click('Start / repeat lookup'); b.click('Start / repeat lookup'); await a.flush();
  assert.match(a.container.textContent, /Loading hub URL/);
  await a.flush(() => a.calls[0].resolve({ value: JSON.stringify({ url: 'https://tenant.test/HUB-A' }) }));
  assert.match(a.container.textContent, /HUB-A/); assert.doesNotMatch(b.container.textContent, /HUB-A/);
  assert.ok(a.observations.some(record => /HUB-A/.test(record.text)), 'Profiler sees valid URL commits, so scope assertions cannot pass by always hiding results');
  await b.flush(() => b.calls[0].reject(new Error('Hub unavailable'))); assert.match(b.container.textContent, /Error: Hub unavailable/);
  const count = a.calls.length; a.render(); assert.equal(a.calls.length, count);
  a.click('Start / repeat lookup'); await a.flush(); assert.equal(a.calls.length, count + 1);
  await a.flush(() => a.calls.at(-1).resolve({ value: '{}' })); assert.match(a.container.textContent, /Hub URL not available yet/); assert.doesNotMatch(a.container.textContent, /HUB-A/);
  a.click('Start / repeat lookup'); await a.flush(); await a.flush(() => a.calls.at(-1).resolve({ value: 'invalid JSON' })); assert.match(a.container.textContent, /Error:/);
  for (const id of [undefined, '', '00000000-0000-0000-0000-000000000000']) {
    a.instance.context.pageContext.legacyPageContext.hubSiteId = id; a.render();
    const before = a.calls.length; a.click('Start / repeat lookup'); await a.flush();
    assert.match(a.container.textContent, /No hub association/); assert.equal(a.calls.length, before);
  }
  const page = a.instance.context.pageContext; page.legacyPageContext.hubSiteId = 'aaaaaaaa-0000-0000-0000-000000000001'; page.legacyPageContext.isHubSite = true; page.webAbsoluteUrl = 'https://tenant.test/CURRENT-HUB'; a.render();
  const before = a.calls.length; a.click('Start / repeat lookup'); await a.flush(); assert.match(a.container.textContent, /CURRENT-HUB/); assert.equal(a.calls.length, before);
});
