import * as React from 'react';
import { createDOMRenderer, makeStyles } from '@griffel/core';
import { RendererProvider, useRenderer_unstable } from '@griffel/react';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme } from '@fluentui/react-theme';
import { useSx } from '__HOOK__';
import { useSx as useOtherSx } from '__OTHER_HOOK__';
import { width, marginInlineStart } from '__DESCRIPTORS__';

const physical=makeStyles({root:{marginLeft:'7px'}});

export function run({ReactDOM, act, assert}) {
  const renderer=createDOMRenderer(document), otherRenderer=createDOMRenderer(document);
  const container=document.createElement('div');document.body.appendChild(container);
  const observations={};
  const descriptor=marginInlineStart.px(12);
  function Child({id}) {
    const sx=useSx(), other=useOtherSx();
    const currentRenderer=useRenderer_unstable(), {dir}=useFluent_unstable();
    const external=physical({renderer:currentRenderer,dir}).root;
    observations[id]={first:sx(width.full,descriptor,external),second:other(width.full,descriptor,external)};
    return <span className={observations[id].first}/>;
  }
  function view(dir) {
    return <><RendererProvider renderer={renderer}><FluentProvider dir={dir} theme={webLightTheme}><Child id="main" /></FluentProvider></RendererProvider><RendererProvider renderer={otherRenderer}><FluentProvider dir={dir} theme={webLightTheme}><Child id="other" /></FluentProvider></RendererProvider></>;
  }
  const rules=renderer=>Object.values(renderer.stylesheets).flatMap(sheet=>Array.from(sheet.element.sheet.cssRules).map(rule=>rule.cssText));
  act(()=>ReactDOM.render(view('ltr'),container));
  assert.equal(observations.main.first,observations.main.second);
  assert.equal(observations.other.first,observations.other.second);
  assert.equal(observations.main.first,observations.other.first);
  const ltr=observations.main.first, before=rules(renderer);
  assert.ok(before.length>0);assert.ok(before.some(rule=>/--apvee-sx-width-base:\s*100%/.test(rule)),'Missing width descriptor CSS variable binding');
  assert.deepEqual(before,rules(otherRenderer));
  act(()=>ReactDOM.render(view('ltr'),container));
  assert.deepEqual(rules(renderer),before);assert.equal(observations.main.first,ltr);
  act(()=>ReactDOM.render(view('rtl'),container));
  assert.equal(observations.main.first,observations.main.second);
  assert.deepEqual(rules(renderer),rules(otherRenderer));
  assert.ok(Object.keys(renderer.insertionCache).length>0);
  assert.notEqual(renderer.insertionCache,otherRenderer.insertionCache);
  const rtl=observations.main.first;assert.notEqual(rtl,ltr);
  act(()=>ReactDOM.unmountComponentAtNode(container));container.remove();
  return {assertions:['mixed useSx imports share real Griffel/Fluent contexts','class and CSS rule equivalence','rerender deduplicates rules','LTR/RTL contexts and independent renderers'],ltr,rtl,ruleCount:rules(renderer).length};
}
