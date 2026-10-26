const test = require('node:test'); const assert = require('node:assert/strict');
const { createHarness, press, paste } = require('./fixtures/editable-text/editor-harness.cjs');
for (const blocks of [['p'], ['h1','h2','h3','h4'], ['h2','h4']]) test(`real engine enforces ${blocks} for commands, paste, reset and shortcuts`, t => {
  const h = createHarness(); t.after(() => h.close()); const policy = h.policy(blocks); const f = h.mount({ policy }); const e = f.editor;
  const allowed = () => { const b = e.getJSON().content[0]; assert.ok(blocks.includes(b.type === 'paragraph' ? 'p' : `h${b.attrs.level}`)); };
  for (const key of ['0','1','2','3','4']) { press(h,e,key,{ctrlKey:true,altKey:true}); allowed(); }
  for (const [html,text] of [['','plain\nnext'],['<h1><strong>A</strong></h1><p>B</p>',''],['<p>Plain</p>','']]) { paste(h,e,html,text); allowed(); }
  for (const action of [() => e.commands.setParagraph(), () => e.commands.setHeading({level:1}), () => e.commands.clearContent(), () => e.commands.undo(), () => e.commands.redo()]) { h.act(action); allowed(); }
  for (const doc of f.calls) { const b = doc.type === 'paragraph' ? 'p' : `h${doc.level}`; assert.ok(blocks.includes(b)); }
});
test('headings-only toggle cannot escape to paragraph and deliberate empty allowed heading retains stored marks', t => {
  const h = createHarness(); t.after(() => h.close()); const f = h.mount({ policy: h.policy(['h2','h4']) }); const e = f.editor;
  h.act(() => { e.commands.clearContent(); e.commands.setHeading({level:4}); e.commands.setMark('bold'); });
  press(h,e,'4',{ctrlKey:true,altKey:true}); assert.equal(e.getJSON().content[0].attrs.level,4);
  h.act(() => e.commands.insertContent('X')); assert.equal(h.serialize(f.calls.at(-1)), '<h4><strong>X</strong></h4>');
});
test('superset schema normalizes excluded setNode before callback and isolates two policies', t => {
  const h = createHarness(); t.after(() => h.close()); const a = h.mount({policy:h.policy(['p'])}); const b = h.mount({policy:h.policy(['h2','h4'])});
  h.act(() => a.editor.commands.setNode('heading',{level:4})); assert.equal(a.editor.getJSON().content[0].type,'paragraph');
  h.act(() => b.editor.commands.setNode('paragraph')); assert.equal(b.editor.getJSON().content[0].attrs.level,2);
  assert.equal(a.editor.getJSON().content[0].type,'paragraph');
});
test('discontinuous heading input rules match exact permitted levels and keep excluded markers literal',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({policy:h.policy(['h2','h4'])}); const e=f.editor;
 for(const [marker,level] of [['# ',undefined],['## ',2],['### ',undefined],['#### ',4]]) {
  h.act(()=>e.commands.setContent({type:'doc',content:[{type:'heading',attrs:{level:2}}]}));
  for(const text of marker) h.act(()=>{
   const from=e.state.selection.from,to=e.state.selection.to;
   const handled=e.view.someProp('handleTextInput',handler=>handler(e.view,from,to,text));
   if(!handled) e.commands.insertContent(text);
  });
  if(level) { assert.equal(e.getJSON().content[0].attrs.level,level); assert.equal(e.getText(),''); }
  else assert.equal(e.getText(),marker);
 }
});
test('all-block shortcuts retain native heading toggle and reset applies an explicit default to an already empty editor',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({policy:h.policy(undefined,'h4')}); const e=f.editor;
 press(h,e,'2',{ctrlKey:true,altKey:true}); assert.equal(e.getJSON().content[0].attrs.level,2);
 press(h,e,'2',{ctrlKey:true,altKey:true}); assert.equal(e.getJSON().content[0].type,'paragraph');
 h.act(()=>e.commands.clearContent()); assert.equal(e.getJSON().content[0].attrs.level,4);
 h.act(()=>e.commands.setParagraph()); assert.equal(e.getJSON().content[0].type,'paragraph');
 h.act(()=>e.commands.clearContent()); assert.equal(e.getJSON().content[0].attrs.level,4);
});
test('HTML paste rejects an import whose policy projection exceeds canonical byte budget',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({policy:h.policy(['h2'])}); const e=f.editor;
 paste(h,e,'<p>'+'x'.repeat(65529)+'</p>');
 assert.equal(e.getText(),'initial'); assert.equal(f.calls.length,0);
});
test('one raw valid excluded-block transaction publishes only the final compliant document',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({policy:h.policy(['h2','h4'])}); const e=f.editor;
 h.act(()=>e.commands.setContent('<h1><strong>Changed</strong></h1>'));
 assert.equal(f.calls.length,1); assert.equal(h.serialize(f.calls[0]),'<h2><strong>Changed</strong></h2>');
 h.act(()=>e.commands.undo()); assert.equal(e.getText(),'initial');
 h.act(()=>e.commands.redo()); assert.equal(e.getHTML(),'<h2><strong>Changed</strong></h2>');
});
test('headings-only toggle reads the current chained transaction instead of stale editor state',t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({document:{type:'heading',level:4,content:[{type:'text',text:'initial'}]},policy:h.policy(['h2','h4'])});
 h.act(()=>f.editor.chain().setHeading({level:2}).toggleHeading({level:4}).run());
 assert.equal(f.editor.getJSON().content[0].attrs.level,4);
});
