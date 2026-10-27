import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Button, Input, Label, Link, MessageBar, MessageBarActions, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
/**
 * Example 3: useSPFxPnPSearch - Basic Search
 */
export const PnPSearchBasicDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const [executedTarget, setExecutedTarget] = React.useState<string>();
    const [searchText, setSearchText] = React.useState('');
    const { search, results, totalResults, loading, error, clearError, } = useSPFxPnPSearch<{
        Title: string;
        Path: string;
        FileType: string;
    }>({
        pageSize: 10
    });
    const handleSearch = React.useCallback(async () => {
        if (!searchText.trim())
            return;
        try {
            setExecutedTarget(searchText);
            await search(searchText);
        }
        catch (err) {
            console.error('Search error:', err);
        }
    }, [searchText, search]);
    return (<DemoCard title="Example 1: Basic Search" description="Search SharePoint using a text query." iconName="Search" controls={<>{selector?.(false)}
    <div className={demoStyles.actions}>
      <Input aria-label={"Search SharePoint..."} value={searchText} onChange={(_, data) => setSearchText(data.value ?? '')} placeholder="Search SharePoint..." onKeyDown={(e) => {
                if (e.key === 'Enter') {
                    handleSearch().catch(err => console.error('Search error:', err));
                }
            }}/>

    </div>
  </>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={handleSearch} disabled={!searchText.trim() || loading}>{loading ? 'Searching...' : 'Search'}</Button>
  </>} diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      Simple text search. Try: &quot;document&quot;, &quot;report&quot;, &quot;ContentType:Document&quot;
    </Label>
  </>}>

    {loading ? <Label>Loading...</Label> : executedTarget !== searchText ? <Label>Run the selected search.</Label> : <>
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
        Found {totalResults} result{totalResults !== 1 ? 's' : ''}
      </MessageBarBody>
    </MessageBar>)}

    {results.length > 0 ? (<div className={demoStyles.panel}>
      {results.map((result) => (<div className={demoStyles.panel} key={result.id}>
        <div className={demoStyles.observations}>{result.data.Title || '(No Title)'}</div>
        <Link className={demoStyles.observations} href={result.data.Path} target="_blank" rel="noopener noreferrer">
          {result.data.Path}
        </Link>
        {result.data.FileType && <Label>Type: {result.data.FileType}</Label>}
      </div>))}
    </div>) : totalResults === 0 ? (<MessageBar>
      <MessageBarBody>No results found. Try a different query.</MessageBarBody>
    </MessageBar>) : null}

    </>}

  </DemoCard>);
};
