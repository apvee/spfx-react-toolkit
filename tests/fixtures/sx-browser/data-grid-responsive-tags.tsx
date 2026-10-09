import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { createDataGrid, useSx, flex, gap, width } from '@apvee/spfx-react-toolkit';

interface TagRow { readonly id: string; readonly tags: readonly ('design' | 'review' | 'release')[]; }
const Grid = createDataGrid<TagRow>();
const items: readonly TagRow[] = [{ id: 'responsive', tags: ['design', 'review', 'release'] }];
const values = { design: { label: 'Design' }, review: { label: 'Review' }, release: { label: 'Release' } };

/** Isolate genuine native responsive overflow; no configured maxVisible or mock measurement. */
export function ResponsiveTagsFixture(): React.ReactElement {
  const sx = useSx();
  const [size, setSize] = React.useState(300);
  return <FluentProvider theme={webLightTheme}><section id="dg-responsive-tags-fixture" className={sx(flex.column, gap.small)}>
    <label>Responsive Tags fixture width <select id="dg-responsive-tags-width" value={size} onChange={event => setSize(Number(event.target.value))}><option value="300">300px</option><option value="100">100px</option></select></label>
    <div id="dg-responsive-tags-frame" className={sx(width.px(size))}>
      <Grid id="dg-responsive-tags-grid" aria-label="Responsive Tags fixture" items={items} getRowId={item => item.id}
        resizableColumns resizableColumnsOptions={{ autoFitColumns: true }} columnSizingOptions={{ tags: { minWidth: 60 } }}>
        <Grid.TagsColumn id="tags" header="Responsive Tags" field="tags" values={values} minWidth={60} width={300} />
      </Grid>
    </div>
  </section></FluentProvider>;
}
