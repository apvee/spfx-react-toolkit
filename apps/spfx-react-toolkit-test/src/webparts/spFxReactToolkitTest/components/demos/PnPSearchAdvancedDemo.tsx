import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Button, Label, Link, MessageBar, MessageBarActions, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPSearch, SearchVerticals } from '@apvee/spfx-react-toolkit';
/**
 * Example 4: Advanced Search with Builder and Verticals
 */
export const PnPSearchAdvancedDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const [executedTarget, setExecutedTarget] = React.useState<string>();
    const [selectedVertical, setSelectedVertical] = React.useState<string>('All');
    const { search, results, totalResults, loading, loadMore, hasMore, error, clearError, } = useSPFxPnPSearch<{
        Title: string;
        Path: string;
        Author: string;
        LastModifiedTime: string;
        FileType: string;
    }>({
        pageSize: 5,
        selectProperties: ['Title', 'Path', 'Author', 'LastModifiedTime', 'FileType']
    });
    const getVerticalSourceId = (vertical: string): string | undefined => {
        switch (vertical) {
            case 'People': return SearchVerticals.People;
            case 'Documents': return SearchVerticals.Documents;
            case 'Pages': return SearchVerticals.Pages;
            case 'Videos': return SearchVerticals.Videos;
            default: return undefined;
        }
    };
    const handleSearch = React.useCallback((query: string) => {
        setExecutedTarget(selectedVertical);
        search((builder) => {
            let result = builder.text(query).sortList({ Property: 'LastModifiedTime', Direction: 1 });
            const sourceId = getVerticalSourceId(selectedVertical);
            if (sourceId)
                result = result.sourceId(sourceId);
            return result;
        }).catch(err => console.error('Search error:', err));
    }, [search, selectedVertical]);
    return (<DemoCard title="Example 2: Advanced Search with Builder" description="Search a selected vertical with builder queries and pagination." iconName="SearchAndApps" controls={<>{selector?.(false)}
    <div className={demoStyles.actions}>
      {['All', 'Documents', 'Pages', 'People', 'Videos'].map((v) => (<Button key={v} onClick={() => setSelectedVertical(v)} appearance={selectedVertical === v ? "primary" : "secondary"}>{v}</Button>))}
    </div>

  </>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={() => handleSearch('*')} disabled={loading}>{"Search All"}</Button>
    <Button onClick={() => handleSearch('filetype:docx OR filetype:pdf')} disabled={loading}>{"Search Documents"}</Button>
  </>} diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      Builder API with verticals (People, Documents, Pages, Videos) and pagination.
    </Label>
  </>}>

    {loading ? <Label>Loading...</Label> : executedTarget !== selectedVertical ? <Label>Run the selected search.</Label> : <>
    {error && (<MessageBar intent="error">
      <MessageBarBody>
        {error.message}
      </MessageBarBody>
      <MessageBarActions>
        <Button aria-label="Dismiss message" onClick={clearError}>Dismiss</Button>
      </MessageBarActions>
    </MessageBar>)}

    {totalResults !== undefined && (<MessageBar intent="info">
      <MessageBarBody>
        Found {totalResults} result{totalResults !== 1 ? 's' : ''} in &quot;{selectedVertical}&quot;
      </MessageBarBody>
    </MessageBar>)}

    {results.length > 0 && (<div className={demoStyles.panel}>
      {results.map((result) => (<div className={demoStyles.panel} key={result.id}>
        <div className={demoStyles.observations}>{result.data.Title || '(No Title)'}</div>
        <Label>{result.data.Author} - {result.data.LastModifiedTime ? new Date(result.data.LastModifiedTime).toLocaleDateString() : 'N/A'}</Label>
        <Link className={demoStyles.observations} href={result.data.Path} target="_blank" rel="noopener noreferrer">
          {result.data.Path}
        </Link>
      </div>))}
      {hasMore && (<Button appearance="primary" onClick={loadMore} disabled={loading}>{loading ? 'Loading...' : 'Load More'}</Button>)}
    </div>)}

    </>}

  </DemoCard>);
};
