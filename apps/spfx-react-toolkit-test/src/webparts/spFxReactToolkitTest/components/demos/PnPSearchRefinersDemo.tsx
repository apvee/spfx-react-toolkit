import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Button, Field, Label, Link, MessageBar, MessageBarActions, MessageBarBody, Select } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
/**
 * Example 5: Faceted Search with Refiners
 */
export const PnPSearchRefinersDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const [refinerOperator, setRefinerOperator] = React.useState<'and' | 'or'>('or');
    const { search, results, refiners, totalResults, loading, applyRefiner, error, clearError, } = useSPFxPnPSearch<{
        Title: string;
        Path: string;
        FileType: string;
        Author: string;
    }>({
        pageSize: 10,
        selectProperties: ['Title', 'Path', 'FileType', 'Author'],
        // DisplayAuthor preserves the names used by the author refinement buckets.
        refiners: 'FileType,DisplayAuthor,Size,Write',
        refinerOperator
    });
    const handleSearch = React.useCallback(() => {
        search('*').catch(err => console.error('Search error:', err));
    }, [search]);
    return (<DemoCard title="Example 3: Faceted Search with Refiners" description="Search content and filter results using refinement buckets." iconName="FilterSettings" controls={<>{selector?.(false)}
    <Field label="Match values within each facet">
      <Select value={refinerOperator} disabled={loading} onChange={event => {
                setRefinerOperator(event.target.value === 'or' ? 'or' : 'and');
            }}>
        <option value="or">Any selected value (OR)</option>
        <option value="and">All selected values (AND)</option>
      </Select>
    </Field>
    <Label>Different facets combine with AND. Changing this choice resets the search.</Label>

  </>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={handleSearch} disabled={loading}>{loading ? 'Loading...' : 'Search All Content'}</Button>
  </>} diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      Faceted search with FileType, DisplayAuthor, Size and Write refiners. Click to filter/toggle.
    </Label>
  </>}>

    {error && (<MessageBar intent="error">
      <MessageBarBody>
        {error.message}
      </MessageBarBody>
      <MessageBarActions>
        <Button aria-label="Dismiss message" onClick={clearError}>Dismiss</Button>
      </MessageBarActions>
    </MessageBar>)}

    {!loading && !error && totalResults !== undefined && (<Label>
      <DemoIcon name={"SearchIssue"}/>
      Found {totalResults} results
    </Label>)}

    {loading && <Label>Loading...</Label>}
    <div className={demoStyles.actions}>
      {!loading && !error && refiners.length > 0 && (<div className={demoStyles.panel}>
        <Label className={demoStyles.observations}>Filters</Label>
        {refiners.map((refiner) => (<div className={demoStyles.panel} key={refiner.name}>
          <Label>{refiner.name}</Label>
          {refiner.entries.slice(0, 5).map((entry) => (<Button key={entry.value} disabled={loading || !entry.token} onClick={() => applyRefiner(refiner.name, entry.token).catch(err => console.error('Refiner error:', err))}>{`${entry.value} (${entry.count})`}</Button>))}
        </div>))}
      </div>)}

      <div className={demoStyles.panel}>
        {!loading && !error && results.length > 0 ? (results.map((result) => (<div className={demoStyles.panel} key={result.id}>
          <div className={demoStyles.observations}>{result.data.Title || '(No Title)'}</div>
          <Label>{result.data.FileType} | {result.data.Author}</Label>
          <Link className={demoStyles.observations} href={result.data.Path} target="_blank" rel="noopener noreferrer">
            {result.data.Path}
          </Link>
        </div>))) : (<Label>Click &quot;Search All Content&quot; to see results.</Label>)}
      </div>
    </div>

  </DemoCard>);
};
