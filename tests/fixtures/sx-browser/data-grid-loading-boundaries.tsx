import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { createDataGrid, useSx, width } from '@apvee/spfx-react-toolkit';

interface Person { readonly id: string; readonly name: string; }
interface LoadingRow {
  readonly id: string;
  readonly title: string;
  readonly person: Person;
  readonly rating: number;
  readonly people: readonly Person[];
  readonly tags: readonly 'tag'[];
}
const Grid = createDataGrid<LoadingRow>();
const items: readonly LoadingRow[] = [];
const tags = { tag: { label: 'Tag' } };
const FallbackGrid = createDataGrid<Pick<LoadingRow, 'id' | 'person'>>();
const fallbackItems: readonly Pick<LoadingRow, 'id' | 'person'>[] = [
  { id: 'actual', person: { id: 'owner', name: 'Visible fallback owner' } },
];

/** Tiny compound columns exercise bounded templates and AutoSkeleton's size fallback. */
export function LoadingBoundariesFixture(): React.ReactElement {
  const sx = useSx();
  return <FluentProvider theme={webLightTheme}><section id="dg-loading-boundaries" className={sx(width.px(230))}>
    <Grid id="dg-loading-boundaries-grid" aria-label="Narrow loading template fixture" items={items}
      loading getRowId={item => item.id}>
      <Grid.PersonColumn id="person" header="P" field="person" width={20} minWidth={20}
        getName={person => person.name} />
      <Grid.RatingColumn id="rating" header="R" field="rating" max={5} width={20} minWidth={20} />
      <Grid.PeopleColumn id="people" header="Pe" field="people" width={20} minWidth={20}
        getName={person => person.name} getPersonId={person => person.id} />
      <Grid.TagsColumn id="tags" header="T" field="tags" values={tags} width={20} minWidth={20} />
      <Grid.TextColumn id="title" header="Text" field="title" width={150} minWidth={150} />
    </Grid>
  </section><section id="dg-loading-unmeasurable" className={sx(width.px(20))}>
    <FallbackGrid id="dg-loading-unmeasurable-grid" aria-label="Actual narrow person fallback" items={fallbackItems}
      loading getRowId={item => item.id}>
      {/* Native cells use content-box: 4px content plus 8px padding on each side is 20px overall. */}
      <FallbackGrid.PersonColumn id="person" header="P" field="person" width={4} minWidth={4}
        getName={person => person.name} />
    </FallbackGrid>
  </section></FluentProvider>;
}
