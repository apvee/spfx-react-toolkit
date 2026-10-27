const test=require('node:test'); const assert=require('node:assert/strict'); const {createHarness}=require('./fixtures/editable-text/editor-harness.cjs');
for(const blocks of [['p'],['h2','h4']]) test(`direct block selector respects ${blocks}`,async t=>{
 const h=createHarness(); t.after(()=>h.close()); const f=h.mount({policy:h.policy(blocks)}); const e=f.editor;
 h.act(()=>e.view.dom.dispatchEvent(new h.window.FocusEvent('focusin',{bubbles:true}))); await h.settle();
 const button=document.querySelector('[aria-label="Heading"]');
 if(blocks.length===1) assert.equal(button,null); else { assert.ok(button); h.act(()=>button.click());
 assert.deepEqual([...document.querySelectorAll('[role="option"]')].map(n=>n.textContent.trim()),['Heading 2','Heading 4']); }
});
for(const blocks of [['p'],['h2','h4']]) test(`overflow block menu respects ${blocks}`,async t=>{
 const h=createHarness(); t.after(()=>h.close());
 const prototype=h.window.HTMLElement.prototype; const descriptor=Object.getOwnPropertyDescriptor(prototype,'clientWidth');
 Object.defineProperty(prototype,'clientWidth',{configurable:true,get(){return this.getAttribute('role')==='toolbar'?96:descriptor.get.call(this);}});
 t.after(()=>Object.defineProperty(prototype,'clientWidth',descriptor));
 const f=h.mount({policy:h.policy(blocks)}); const e=f.editor;
 h.act(()=>e.view.dom.dispatchEvent(new h.window.FocusEvent('focusin',{bubbles:true}))); await h.settle();
 const more=document.querySelector('[aria-label="More formatting"]')||document.querySelector('[aria-label="More text formatting"]'); assert.ok(more,'native manager mounts More');
 h.act(()=>more.click()); const heading=[...document.querySelectorAll('[role="menuitem"]')].find(n=>n.textContent.trim()==='Heading');
 if(blocks.length===1) assert.equal(heading,undefined); else { assert.ok(heading); h.act(()=>heading.click());
 assert.deepEqual([...document.querySelectorAll('[role="menuitemradio"]')].map(n=>n.textContent.trim()),['Heading 2','Heading 4']);
 h.act(()=>[...document.querySelectorAll('[role="menuitemradio"]')].find(n=>n.textContent.trim()==='Heading 4').click()); assert.equal(e.getJSON().content[0].attrs.level,4); }
});
