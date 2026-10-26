const test = require('node:test'); const assert = require('node:assert/strict');
const { createHarness } = require('./fixtures/editable-text/editor-harness.cjs');
test('membership changes operate on live draft, retain view/range/marks and reset history without publication', t => {
  const h=createHarness(); t.after(()=>h.close()); const f=h.mount(); const e=f.editor, view=e.view;
  h.act(()=>{ e.commands.setTextSelection({from:2,to:5}); e.commands.toggleBold(); }); const count=f.calls.length;
  f.render({policy:h.policy(['h2','h4'])}); assert.equal(e.view,view); assert.equal(e.getHTML(),'<h2>i<strong>nit</strong>ial</h2>');
  assert.equal(e.state.selection.from,2); assert.equal(e.state.selection.to,5); assert.equal(f.calls.length,count);
  h.act(()=>e.commands.undo()); assert.equal(e.getHTML(),'<h2>i<strong>nit</strong>ial</h2>');
});
test('equivalent membership and default-only updates retain undo and allowed nonempty block', t => {
  const h=createHarness(); t.after(()=>h.close()); const f=h.mount({policy:h.policy(['p','h2'])}); const e=f.editor;
  h.act(()=>e.commands.insertContent('X')); f.render({policy:h.policy(['h2','p','p'],'h2')});
  assert.equal(e.getHTML(),'<p>Xinitial</p>'); h.act(()=>e.commands.undo()); assert.equal(e.getHTML(),'<p>initial</p>');
});
test('oversized policy-only conversion preserves draft and error without callback', t=>{
  const h=createHarness(); t.after(()=>h.close()); const f=h.mount(); const e=f.editor;
  h.act(()=>e.commands.setContent({type:'doc',content:[{type:'paragraph',content:[{type:'text',text:'&'.repeat(14000)}]}]}));
  assert.ok(f.container.textContent.includes('too large')); f.render({policy:h.policy(['h4'])});
  assert.equal(e.getText().length,14000); assert.equal(e.getJSON().content[0].attrs.level,4); assert.equal(f.calls.length,0); assert.ok(f.container.textContent.includes('too large'));
});
for (const local of [false,true]) test(`IME A→B→C uses latest policy with ${local?'one':'zero'} final callback`,async t=>{
  const h=createHarness(); t.after(()=>h.close()); const f=h.mount(); const e=f.editor;
  const event=type=>h.act(()=>e.view.dom.dispatchEvent(new h.window.CompositionEvent(type,{bubbles:true})));
  event('compositionstart'); f.render({policy:h.policy(['h2'])}); f.render({policy:h.policy(['h4'])});
  if(local) h.act(()=>e.commands.insertContent('IME ')); assert.equal(f.calls.length,0);
  event('compositionend'); await h.settle(); assert.equal(e.getJSON().content[0].attrs.level,4); assert.equal(f.calls.length,local?1:0);
  assert.equal(e.getText(),local?'IME initial':'initial');
});
test('canonical source echo retains undo; changed source with same props document replaces live draft',t=>{
  const h=createHarness(); t.after(()=>h.close()); const f=h.mount(); const e=f.editor;
  h.act(()=>e.commands.insertContent('X')); const doc=f.calls.at(-1); f.render({document:doc,sourceKey:'echo'});
  h.act(()=>e.commands.undo()); assert.equal(e.getText(),'initial');
  f.render({document:doc,sourceKey:'reset'}); assert.equal(e.getText(),'Xinitial');
});
test('membership reset retains stored marks on an empty allowed heading before typing',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({document:{type:'heading',level:4,content:[]},policy:h.policy(['h2','h4'],'h4')}); const e=f.editor;
 h.act(()=>e.commands.setMark('bold')); f.render({policy:h.policy(['h1','h2','h4'],'h2')});
 assert.equal(e.getJSON().content[0].attrs.level,4); assert.ok(e.state.storedMarks.some(mark=>mark.type.name==='bold'));
 h.act(()=>e.commands.insertContent('X'));
 assert.equal(h.serialize(f.calls.at(-1)),'<h4><strong>X</strong></h4>');
});
test('latest policy wins final microtask input and unmount flush; policy-only IME preserves undo until completion',async t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount(); const e=f.editor;
 const event=type=>h.act(()=>e.view.dom.dispatchEvent(new h.window.CompositionEvent(type,{bubbles:true})));
 event('compositionstart'); f.render({policy:h.policy(['h2'])});
 assert.equal(e.getJSON().content[0].type,'paragraph'); event('compositionend');
 queueMicrotask(()=>h.act(()=>e.commands.insertContent('final '))); await h.settle();
 assert.equal(f.calls.length,1); assert.equal(h.serialize(f.calls[0]),'<h2>final initial</h2>');
 event('compositionstart'); h.act(()=>e.commands.insertContent('last ')); f.render({policy:h.policy(['h4'])}); f.unmount();
 assert.equal(f.calls.length,2); assert.equal(h.serialize(f.calls[1]),'<h4>final last initial</h4>');
});
test('explicit default-only change projects an empty live draft without callback or resetting undo',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount(); const e=f.editor;
 h.act(()=>e.commands.setTextSelection({from:1,to:8})); h.act(()=>e.commands.deleteSelection());
 const count=f.calls.length,view=e.view; assert.equal(e.getText(),'');
 f.render({policy:h.policy(undefined,'h2')}); assert.equal(e.view,view); assert.equal(e.getJSON().content[0].attrs.level,2); assert.equal(f.calls.length,count);
 h.act(()=>e.commands.undo()); assert.equal(e.getText(),'initial');
});
