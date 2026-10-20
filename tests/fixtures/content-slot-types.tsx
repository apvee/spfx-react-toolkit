import * as React from 'react';
import type { ContentSlot, AlternativeContentSlot } from '../../packages/spfx-react-toolkit/src/controls/content-slot.types.internal';
declare const legacy: React.ReactNode;
const accepted: ContentSlot<'span'>[] = ['Text', 0, false, true, undefined, null, legacy, <></>, ['A'], new Set(['A']),
  { as: 'span', ref: React.createRef<HTMLSpanElement>(), children: (Component, props) => <Component {...props}>Callback</Component> },
  { className: 'custom', style: { color: 'red' }, 'aria-label': 'Name', 'data-observed': 'yes', onClick: event => event.currentTarget.focus() }];
const owned: ContentSlot<'span', 'id' | 'aria-labelledby'> = { children: 'Safe' };
const logo: AlternativeContentSlot<'div', 'img'> = { as: 'img', src: '/logo.png', alt: 'Logo', ref: React.createRef<HTMLImageElement>() };
// @ts-expect-error fixed as cannot change heading
const wrongTag: ContentSlot<'span'> = { as: 'h1' };
// @ts-expect-error raw HTML unsupported
const raw: ContentSlot<'span'> = { dangerouslySetInnerHTML: { __html: 'unsafe' } };
// @ts-expect-error owner controls id
const wrongId: ContentSlot<'span', 'id'> = { id: 'unsafe' };
// @ts-expect-error unknown props unsupported
const unknown: ContentSlot<'span'> = { madeUp: true };
// @ts-expect-error a direct callback is not shorthand
const callback: ContentSlot<'span'> = () => 'No';
// @ts-expect-error bounded alternate tags
const wrongLogo: AlternativeContentSlot<'div', 'img'> = { as: 'a', href: '/' };
void [accepted, owned, logo, wrongTag, raw, wrongId, unknown, callback, wrongLogo];


// String-only class composition must support lightweight consumers without the toolkit styles engine.
import { mergeClasses } from '@griffel/react';
import { renderContentSlot } from '../../packages/spfx-react-toolkit/src/controls/content-slot.internal';
import { useSx } from '../../packages/spfx-react-toolkit/src/hooks/useSx';
declare const sxComposer: ReturnType<typeof useSx>;
renderContentSlot({ children: 'Lightweight' }, 'span', mergeClasses, { shorthand: 'preserve' });
renderContentSlot({ children: 'Themed' }, 'span', sxComposer, { shorthand: 'preserve' });
