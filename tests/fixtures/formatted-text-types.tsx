import type { FormattedTextProps, FormattedTextVariant, FormattedTextSemanticElement } from '../../packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text.types';
import * as typography from '../../packages/spfx-react-toolkit/src/helpers/styles/typography';

const variants: FormattedTextVariant[] = Object.keys(typography) as FormattedTextVariant[];
const elements: FormattedTextSemanticElement[] = ['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
const props: FormattedTextProps[] = variants.map((variant): FormattedTextProps => ({
  variant, as: 'span', maxLines: 2, font: 'monospace', align: 'justify',
  subtle: true, highlight: true, italic: true, underline: true, strikethrough: true,
  markdown: false, truncate: true, wrap: false, className: 'custom',
  style: { WebkitLineClamp: 3 }, id: 'text', title: 'Native title', dir: 'rtl',
  'aria-label': 'Label', children: ['Text', 42], onClick: event => event.currentTarget.focus(),
}));
elements.forEach(as => props.push({ as, maxLines: 1 }));
// @ts-expect-error Unknown typography recipe.
props.push({ variant: 'unknown' });
// @ts-expect-error Restricted semantic element.
props.push({ as: 'div' });
// @ts-expect-error Font families are a closed domain.
props.push({ font: 'serif' });
// @ts-expect-error No separate inline mode.
props.push({ inline: true });
// @ts-expect-error Typography is selected through variant.
props.push({ size: 400 });
// @ts-expect-error Typography is selected through variant.
props.push({ weight: 'bold' });
// @ts-expect-error Element semantics determine natural display.
props.push({ block: true });
// @ts-expect-error Color is supplied through normal consumer styling.
props.push({ color: 'red' });
// @ts-expect-error Children are the only content pathway.
props.push({ dangerouslySetInnerHTML: { __html: '<b>unsafe</b>' } });

import * as React from 'react';
import { FormattedText } from '../../packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text';
const element = <FormattedText as="h2" maxLines={2} ref={React.createRef<HTMLElement>()}>Text</FormattedText>;
void element;
