const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { webLightTheme, webDarkTheme } = require('@fluentui/react-theme');
const helperPath = 'property-pane/pickers/property-pane-picker-theme.internal.ts';
const hookPath = 'property-pane/pickers/usePropertyPanePickerTheme.internal.ts';
const light = { colorNeutralBackground1: '#ffffff', colorNeutralForeground1: '#242424', colorBrandBackground: '#13579b', fontWeightSemibold: '625', spacingHorizontalM: '14px', shadow16: '0 1px 8px #111111', zIndexPopup: '1200' };
const dark = { colorNeutralBackground1: '#292929', colorNeutralForeground1: '#ffffff', colorBrandBackground: '#51aeb2' };
const flush = h => h.act(async () => { await new Promise(resolve => setTimeout(resolve, 45)); });
function fixture() {
  const h = createHarness({ '../../hooks/useSPFxThemeInfo': { useSPFxThemeInfo: () => undefined } });
  const parent = document.createElement('section'), container = document.createElement('div');
  parent.dataset.paneTheme = 'light'; parent.appendChild(container); document.body.appendChild(parent);
  let reads = 0;
  const original = window.getComputedStyle;
  window.getComputedStyle = element => {
    assert.equal(element, container, 'read only the native container, never the inner FluentProvider');
    reads++;
    const palette = element.closest('[data-pane-theme]')?.dataset.paneTheme === 'dark' ? dark : light;
    return { getPropertyValue: key => palette[key.slice(2)] ?? '' };
  };
  return { h, parent, container, get reads() { return reads; }, close() { h.close(); parent.remove(); window.getComputedStyle = original; } };
}
test('reads host tokens outside the control with correct numeric, optional z-index and non-color types', () => {
  const f = fixture();
  try {
    const { readPropertyPanePickerTheme } = f.h.load(helperPath);
    const own = document.createElement('div'); own.dataset.paneTheme = 'dark'; f.container.appendChild(own);
    const theme = readPropertyPanePickerTheme(f.container);
    assert.equal(theme.colorNeutralBackground1, '#ffffff'); assert.equal(theme.colorBrandBackground, '#13579b');
    assert.equal(theme.fontWeightSemibold, 625); assert.equal(theme.zIndexPopup, '1200');
    assert.equal(theme.spacingHorizontalM, '14px'); assert.equal(theme.shadow16, light.shadow16);
    assert.equal(theme.colorNeutralBackground2, webLightTheme.colorNeutralBackground2);
    f.parent.dataset.paneTheme = 'dark';
    assert.equal(readPropertyPanePickerTheme(f.container).colorNeutralBackground2, webDarkTheme.colorNeutralBackground2);
  } finally { f.close(); }
});
test('detached, absent host tokens and invalid numeric tokens do not read the section theme', () => {
  const f = fixture();
  try {
    const { readPropertyPanePickerTheme } = f.h.load(helperPath);
    f.parent.remove(); assert.equal(readPropertyPanePickerTheme(f.container), undefined); assert.equal(f.reads, 0);
    document.body.appendChild(f.parent);
    const original = window.getComputedStyle;
    window.getComputedStyle = element => { assert.equal(element, f.container); return { getPropertyValue: () => '' }; };
    assert.equal(readPropertyPanePickerTheme(f.container), undefined);
    window.getComputedStyle = element => { const style=original(element); return { getPropertyValue: key => key === '--fontWeightSemibold' ? 'invalid' : style.getPropertyValue(key) }; };
    assert.equal(readPropertyPanePickerTheme(f.container).fontWeightSemibold, webLightTheme.fontWeightSemibold);
  } finally { f.close(); }
});
test('uses the owning document window for theme reads', () => {
  const f=fixture();const frame=document.createElement('iframe');document.body.appendChild(frame);
  try {
    const {readPropertyPanePickerTheme}=f.h.load(helperPath);const container=frame.contentDocument.createElement('div');frame.contentDocument.body.appendChild(container);
    frame.contentWindow.getComputedStyle=element=>{assert.equal(element,container);return {getPropertyValue:key=>dark[key.slice(2)]??''};};
    assert.equal(readPropertyPanePickerTheme(container).colorNeutralBackground1,'#292929');assert.equal(f.reads,0);
  } finally {frame.remove();f.close();}
});
test('theme changes and reparenting preserve equivalent snapshots, then clean up on disposal', async () => {
  const f=fixture();let mounted;
  try {
    const {usePropertyPanePickerTheme}=f.h.load(hookPath);
    mounted=f.h.mount(()=>usePropertyPanePickerTheme(f.container));const initial=mounted.current;
    assert.equal(initial.colorNeutralBackground1,'#ffffff');
    f.parent.setAttribute('class','cosmetic');await flush(f.h);assert.equal(mounted.current,initial);
    f.parent.dataset.paneTheme='dark';f.parent.className='dark';await flush(f.h);assert.equal(mounted.current.colorNeutralBackground1,'#292929');
    const changed=mounted.current;
    f.container.appendChild(document.createElement('span'));await flush(f.h);assert.equal(mounted.current,changed);
    const next=document.createElement('section');next.dataset.paneTheme='light';document.body.appendChild(next);next.appendChild(f.container);
    await flush(f.h);assert.equal(mounted.current.colorNeutralBackground1,'#ffffff');
    const before=f.reads;f.parent.className='old detached source';await flush(f.h);assert.equal(f.reads,before);
    mounted.unmount();mounted=undefined;const disposed=f.reads;next.className='after disposal';await flush(f.h);assert.equal(f.reads,disposed);next.remove();
  } finally {mounted?.unmount();f.close();}
});
test('detached onRender connects later and never restores a dark section fallback', async () => {
  const f=fixture();let mounted;
  try {
    const {usePropertyPanePickerTheme}=f.h.load(hookPath);f.parent.remove();
    mounted=f.h.mount(()=>usePropertyPanePickerTheme(f.container));assert.equal(mounted.current,webLightTheme);assert.equal(f.reads,0);
    document.body.appendChild(f.parent);await flush(f.h);assert.equal(mounted.current.colorBrandBackground,'#13579b');
    const theme=mounted.current;const original=window.getComputedStyle;window.getComputedStyle=()=>({getPropertyValue:()=>''});
    f.parent.className='transient';await flush(f.h);assert.equal(mounted.current,theme);window.getComputedStyle=original;
  } finally {mounted?.unmount();f.close();}
});
test('disposed detached fields never read later attachment and new containers do not inherit an old snapshot', async () => {
  const f=fixture();let mounted;
  try {
    const {usePropertyPanePickerTheme}=f.h.load(hookPath);f.parent.remove();
    mounted=f.h.mount(()=>usePropertyPanePickerTheme(f.container));mounted.unmount();mounted=undefined;
    document.body.appendChild(f.parent);await flush(f.h);assert.equal(f.reads,0);
  } finally {mounted?.unmount();f.close();}
});
test('stylesheet replacement and mutation bursts settle without provider feedback or extra state updates', async () => {
  const f=fixture();let mounted;const style=document.createElement('style');
  try {
    const {usePropertyPanePickerTheme}=f.h.load(hookPath);mounted=f.h.mount(()=>usePropertyPanePickerTheme(f.container));
    const initial=mounted.current;const before=f.reads;
    for(let i=0;i<8;i++)f.parent.className='burst-'+i;
    document.head.appendChild(style);await flush(f.h);
    assert.equal(mounted.current,initial);assert.ok(f.reads-before<=2,'coalesce equivalent host/style changes');
    const settled=f.reads;await flush(f.h);assert.equal(f.reads,settled,'no continuing self-feedback');
    f.parent.dataset.paneTheme='dark';const replacement=document.createElement('style');style.replaceWith(replacement);
    await flush(f.h);assert.equal(mounted.current.colorNeutralBackground1,'#292929');
    f.parent.dataset.paneTheme='light';replacement.textContent='.native-theme { color: black; }';await flush(f.h);assert.equal(mounted.current.colorNeutralBackground1,'#ffffff');replacement.remove();
  } finally {mounted?.unmount();style.remove();f.close();}
});
test('real FileFolder theme follows the native bridge rather than the dark web-part fallback', async () => {
  const {createPickerHarness}=require('./fixtures/picker-source-harness.cjs');
  const h=createPickerHarness({
    '../../hooks/useSPFxThemeInfo':{useSPFxThemeInfo:()=>undefined},
    '../../hooks/useSPFxFluent9ThemeInfo':{useSPFxFluent9ThemeInfo:()=>({theme:webDarkTheme})},
    '../../hooks/useSPFxContext':{useSPFxContext:()=>({spfxContext:{pageContext:{cultureInfo:{isRightToLeft:true}}}})},
  });
  const parent=document.createElement('section'),container=document.createElement('div');parent.appendChild(container);document.body.appendChild(parent);
  let palette=light,mode,renders=0;const original=window.getComputedStyle;
  window.getComputedStyle=element=>element===container?{getPropertyValue:key=>palette[key.slice(2)]??''}:original.call(window,element);
  const {PropertyPanePickerHost}=h.load('property-pane/pickers/property-pane-picker-host.internal.tsx');
  const {useFileFolderPickerTheme}=h.load('controls/file-folder-picker/useFileFolderPickerTheme.internal.ts');
  const {createPickerRequestCache}=h.load('controls/pickers/picker-request-cache.internal.ts');
  const runtime=createPickerRequestCache({});
  function Probe(){mode=useFileFolderPickerTheme();renders++;return h.React.createElement('input',{'aria-label':'Retained input'});}
  let mounted;
  try {
    mounted=h.mountComponent(()=>h.React.createElement(PropertyPanePickerHost,{container,runtime},h.React.createElement(Probe)));
    assert.equal(mode,'default');const input=mounted.container.querySelector('input');assert.equal(mounted.container.querySelector('.fui-FluentProvider').dir,'rtl');
    palette=dark;parent.className='native-dark';await flush(h);assert.equal(mode,'dark');assert.equal(mounted.container.querySelector('input'),input);
    const settled=renders;await flush(h);assert.equal(renders,settled);
  } finally {mounted?.unmount();parent.remove();window.getComputedStyle=original;h.close();}
});
test('nonempty unusable anchor colors reject a snapshot rather than overriding a valid native theme', () => {
  const f=fixture();
  try {
    const {readPropertyPanePickerTheme}=f.h.load(helperPath);
    for(const color of ['not-a-color','transparent','currentColor','env(missing)','var(--unknown)','hsl(0 0% 10%)']){
      window.getComputedStyle=()=>({getPropertyValue:key=>key==='--colorNeutralBackground1'?color:key==='--colorNeutralForeground1'?'#ffffff':''});
      assert.equal(readPropertyPanePickerTheme(f.container),undefined,color);
    }
  }finally{f.close();}
});
