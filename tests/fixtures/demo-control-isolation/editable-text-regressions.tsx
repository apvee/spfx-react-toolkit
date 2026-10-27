import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IWebPartShellProperties } from '@apvee/spfx-react-toolkit';
import { EditableText, WebPartShell, useSx, width } from '@apvee/spfx-react-toolkit';

/** Separate regression fixture: the active product tab keeps one subject. */
export const EditableTextIsolationFixture: React.FC<{ instance: BaseClientSideWebPart<IWebPartShellProperties> }> = ({ instance }) => {
  const [paragraph, setParagraph] = React.useState('<h2>Paragraph source</h2>');
  const [heading, setHeading] = React.useState('<p>Heading source</p>');
  const sx = useSx();
  const [narrow, setNarrow] = React.useState(false);
  const [clicks, setClicks] = React.useState(0);
  const [hasRef, setHasRef] = React.useState(false);
  // React 17 callback refs require null when detaching from a DOM node.
  const capture = React.useCallback((element: HTMLDivElement | null): void => { setHasRef(element !== null); }, []);
  // This is a consumer-provided style escape hatch, deliberately forwarded by the control.
  const consumerStyle = React.useMemo(() => ({ borderTopWidth: '3px', borderTopStyle: 'solid' as const }), []);
  return <>
    <button data-editable-isolation="narrow" onClick={() => setNarrow(value => !value)}>Toggle narrow editors</button>
    <div className={sx(width.px(narrow ? 160 : 900))}>
    <EditableText ref={capture} className="editable-consumer-class" style={consumerStyle} onClick={() => setClicks(count => count + 1)}
      data-editable-isolation="paragraph" mode="edit" aria-label="Paragraph editor" strings={{ formattingControls: 'Controlli di formattazione del testo', moreFormatting: 'Altre opzioni di formattazione del testo', heading: 'Intestazione' }} value={paragraph} allowedBlocks={['p']} onChange={data => setParagraph(data.html)} />
    <EditableText data-editable-isolation="headings" mode="edit" aria-label="Heading editor" strings={{ formattingControls: 'Commandes de mise en forme du texte', moreFormatting: 'Autres options de mise en forme du texte', heading: 'Titre', headingLevel: level => `Titre ${level}` }} value={heading} allowedBlocks={['h2', 'h4']} onChange={data => setHeading(data.html)} />
    </div>
    <output data-editable-isolation="native">{JSON.stringify({ hasRef, clicks })}</output>
    <output data-editable-isolation="values">{JSON.stringify({ paragraph, heading })}</output>
    <WebPartShell instance={instance}><p>Independent Shell title shares the common editor.</p></WebPartShell>
  </>;
};
