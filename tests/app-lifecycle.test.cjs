const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {createHarness} = require('./services-test-harness.cjs');

test('Application Customizer recreates a disposed placeholder and ignores former disposal callbacks', async t => {
  const placeholders = [];
  const handlers = new Set();
  let cleanups = 0;
  let harness;
  class ApplicationCustomizer {
    constructor() {
      this.properties = {testMessage:'probe'};
      this.context = {placeholderProvider:{
        changedEvent:{add:(_observer,handler)=>handlers.add(handler),remove:(_observer,handler)=>handlers.delete(handler)},
        tryCreateContent:(_name,options)=>{
          const domElement = document.createElement('div');
          document.body.appendChild(domElement);
          const placeholder = {domElement,dispose:()=>options.onDispose(),notifyDisposal:options.onDispose};
          placeholders.push(placeholder);
          return placeholder;
        }
      }};
    }
  }
  const toolkit = {
    SPFxApplicationCustomizerProvider: props => {
      harness.React.useEffect(()=>()=>{cleanups+=1;},[]);
      return props.children;
    },
    useSPFxInstanceInfo:()=>({id:'extension',kind:'ApplicationCustomizer'}),
    useSPFxPageContext:()=>({web:{title:'Test'}}),
    useSPFxProperties:()=>({properties:{testMessage:'probe'}})
  };
  harness = createHarness({
    '@microsoft/sp-core-library':{Log:{info(){}}},
    '@microsoft/sp-application-base':{BaseApplicationCustomizer:ApplicationCustomizer,PlaceholderName:{Top:0}},
    '@apvee/spfx-react-toolkit':toolkit,
    SpFxReactToolkitTestApplicationCustomizerStrings:{Title:'Probe'}
  });
  t.after(()=>harness.close());
  const file = path.resolve(__dirname,'../apps/spfx-react-toolkit-test/src/extensions/spFxReactToolkitTest/SpFxReactToolkitTestApplicationCustomizer.ts');
  const Customizer = harness.load(file).default;
  const instance = new Customizer();
  await harness.act(async()=>{await instance.onInit();});
  assert.equal(placeholders.length,1);
  assert.match(placeholders[0].domElement.textContent,/Provider active/);
  harness.act(()=>placeholders[0].notifyDisposal());
  assert.equal(cleanups,1);
  harness.act(()=>{for(const handler of handlers)handler();});
  assert.equal(placeholders.length,2,'navigation must recreate the current placeholder');
  harness.act(()=>placeholders[0].notifyDisposal());
  assert.match(placeholders[1].domElement.textContent,/Provider active/,'former disposal must not unmount the replacement');
  harness.act(()=>instance.onDispose());
  assert.equal(cleanups,2);
  assert.equal(handlers.size,0);
});
