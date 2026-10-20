const test = require('node:test');
const assert = require('node:assert/strict');
const { createHarness } = require('./services-test-harness.cjs');
const { ThemeContext_unstable } = require('@fluentui/react-shared-contexts');
const { webDarkTheme, webLightTheme } = require('@fluentui/react-theme');
test('native picker theme follows closest Fluent provider and falls back to host theme', () => {
 const h=createHarness({'../../hooks/useSPFxFluent9ThemeInfo':{useSPFxFluent9ThemeInfo:()=>({theme:webDarkTheme})}});
 const {useFileFolderPickerTheme}=h.load('controls/file-folder-picker/useFileFolderPickerTheme.internal.ts');
 const fallback=h.mount(useFileFolderPickerTheme);
 try {assert.equal(fallback.current,'dark');} finally {fallback.unmount();}
 let observed;const Probe=()=>{observed=useFileFolderPickerTheme();return h.React.createElement('span',undefined,observed);};
 const m=h.mountComponent(()=>h.React.createElement(ThemeContext_unstable.Provider,{value:webDarkTheme},h.React.createElement(ThemeContext_unstable.Provider,{value:webLightTheme},h.React.createElement(Probe))));
 try {assert.equal(observed,'default');}finally{m.unmount();h.close();}
});
