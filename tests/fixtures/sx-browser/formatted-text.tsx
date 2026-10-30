import * as React from 'react';
import { FormattedText } from '../../../packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text';
import type { FormattedTextProps } from '../../../packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text.types';
import { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
import * as foreground from '../../../packages/spfx-react-toolkit/src/helpers/styles/foreground';
import * as background from '../../../packages/spfx-react-toolkit/src/helpers/styles/background';
import * as paddingInline from '../../../packages/spfx-react-toolkit/src/helpers/styles/padding-inline';
import * as borderRadius from '../../../packages/spfx-react-toolkit/src/helpers/styles/border-radius';
import * as typography from '../../../packages/spfx-react-toolkit/src/helpers/styles/typography';

const longText = 'Text wraps across several lines in a bounded preview. العربية نص طويل للاختبار ' +
  'https://example.com/a/very/long/path/with/no/convenient/breaks/abcdefghijklmnopqrstuvwxyz ' +
  'More content makes line clamp one, two, four and five visible. '.repeat(12);
const boundedStyle: React.CSSProperties = { width: 240 };
const clampProps: Pick<FormattedTextProps, 'wrap' | 'truncate' | 'style'> = {
  wrap: false, truncate: true, style: boundedStyle,
};

/** Mounted below the existing fixture FluentProvider, renderer and direction. */
export function FormattedTextFixture(): React.ReactElement {
  const sx = useSx();
  return <section id="ft-fixture" aria-label="FormattedText browser checks" style={{ width: 600, maxWidth: '100%' }}>
    <div><FormattedText id="ft-default">Default inline text</FormattedText></div>
    <FormattedText id="ft-heading" as="h2" variant="title2">Semantic heading</FormattedText>
    <FormattedText id="ft-paragraph" as="p">Natural paragraph layout</FormattedText>
    <div><FormattedText id="ft-subtle" subtle>Subtle role</FormattedText>{' '}
      <span id="ft-subtle-reference" className={sx(foreground.subtle, typography.body1)}>Descriptor subtle role</span></div>
    <div><FormattedText id="ft-code">{'`inline code`'}</FormattedText>{' '}
      <code id="ft-code-reference" className={sx(background.alternative, paddingInline.extraSmall, borderRadius.small)}>Helper code surface</code></div>
    <div><FormattedText id="ft-highlight" highlight>{'Highlighted [link](#ft-fixture) and `code`'}</FormattedText>{' '}
      <span id="ft-mark-reference" className={sx(foreground.primary, paddingInline.extraSmall, borderRadius.small)}>Helper mark foreground/padding/radius</span></div>
    <FormattedText id="ft-clamp-1" maxLines={1} {...clampProps}>{longText}</FormattedText>
    <FormattedText id="ft-clamp-2" maxLines={2} {...clampProps}>{longText}</FormattedText>
    <FormattedText id="ft-clamp-4" maxLines={4} {...clampProps}>{longText}</FormattedText>
    <FormattedText id="ft-clamp-override" maxLines={2} {...clampProps} style={{ width: 240, WebkitLineClamp: 5 }}>{longText}</FormattedText>
    <div><FormattedText id="ft-invalid-clamp" maxLines={1.5}>Invalid clamp keeps natural span display</FormattedText></div>
    <div style={boundedStyle}><FormattedText id="ft-nowrap" wrap={false} style={boundedStyle}>{longText}</FormattedText></div>
    <FormattedText id="ft-truncate" as="p" truncate style={boundedStyle}>{longText}</FormattedText>
    <div>
      <button id="ft-focus-before" type="button">Before links</button>{' '}
      <FormattedText id="ft-focus-markdown">{'[Markdown link](#ft-fixture)'}</FormattedText>{' '}
      <FormattedText><a id="ft-focus-jsx" href="#ft-fixture">Authored JSX link</a></FormattedText>{' '}
      <button id="ft-focus-after" type="button">After links</button>
    </div>
  </section>;
}
