import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { createDOMRenderer } from '@griffel/core';
import { RendererProvider } from '@griffel/react';
import { registerIcons } from '@fluentui/react';
import StylesPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StylesPanel';
registerIcons({ icons: { Color: 'color' } });
ReactDOM.render(<RendererProvider renderer={createDOMRenderer(document)}><StylesPanel /></RendererProvider>, document.getElementById('root'));
window.addEventListener('beforeunload', () => { const root = document.getElementById('root'); if (root) ReactDOM.unmountComponentAtNode(root); });
