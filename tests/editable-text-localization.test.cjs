const test=require('node:test'); const assert=require('node:assert/strict'); const fs=require('node:fs'); const path=require('node:path');
const {createHarness}=require('./fixtures/editable-text/editor-harness.cjs');
test('generic catalogs cover 50 UI locales, aliases, color names and neutral English labels',t=>{
 const h=createHarness(); t.after(()=>h.close()); const file=path.resolve(__dirname,'../packages/spfx-react-toolkit/src/controls/editable-text/localization/editable-text-locales.internal.ts');
 assert.ok(fs.existsSync(file),'generic locale resolver exists'); const a=h.load('controls/editable-text/localization/editable-text-locales.internal.ts');
 assert.equal(Object.keys(a.editableTextLocaleCatalogs).length,50); assert.equal(a.resolveEditableTextLocaleName('IT'),'it-IT'); assert.equal(a.resolveEditableTextLocaleName('unknown'),'en-US');
 for(const s of Object.values(a.editableTextLocaleCatalogs)) { assert.equal(Object.keys(s.colorNames).length,26); assert.ok(!s.headingLevel(4).includes('{')); assert.ok(!s.formattingShortcut('B','Ctrl+B').includes('{')); assert.ok(s.placeholder); }
 const s=a.resolveEditableTextStrings('en'); assert.equal(s.placeholder,'Add text'); assert.ok(!JSON.stringify(s).toLowerCase().includes('title'));
});
test('locale updates refresh labels, templates and placeholder on the existing view without publication',async t=>{
 const h=createHarness(); t.after(()=>h.close()); const {resolveEditableTextStrings:strings}=h.load('controls/editable-text/localization/editable-text-locales.internal.ts');
 const f=h.mount({document:{type:'paragraph',content:[]},strings:strings('it-IT'),placeholder:strings('it-IT').placeholder,ariaLabel:strings('it-IT').editorLabel}); const e=f.editor,view=e.view;
 h.act(()=>e.commands.setMark('bold')); const marks=e.state.storedMarks;
 for(const culture of ['ar-SA','he-IL','de-DE']) { const s=strings(culture); f.render({strings:s,placeholder:s.placeholder,ariaLabel:s.editorLabel}); await h.settle();
 assert.equal(e.view,view); assert.equal(view.dom.getAttribute('aria-label'),s.editorLabel); assert.equal(view.dom.querySelector('[data-placeholder]').getAttribute('data-placeholder'),s.placeholder);
 assert.deepEqual(e.state.storedMarks,marks); assert.equal(f.calls.length,0); }
});
