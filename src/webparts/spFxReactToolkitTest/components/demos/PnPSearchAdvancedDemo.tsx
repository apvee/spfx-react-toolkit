import * as React from 'react';
import {
  Stack,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';
import { useSPFxPnPSearch, SearchVerticals } from '../../../../hooks';

/**
 * Example 4: Advanced Search with Builder and Verticals
 */
export const PnPSearchAdvancedDemo: React.FC = () => {
  const [selectedVertical, setSelectedVertical] = React.useState<string>('All');
  
  const {
    search,
    results,
    totalResults,
    loading,
    loadMore,
    hasMore,
    error,
    clearError,
  } = useSPFxPnPSearch<{ 
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
    search((builder) => {
      let result = builder.text(query).sortList({ Property: 'LastModifiedTime', Direction: 1 });
      const sourceId = getVerticalSourceId(selectedVertical);
      if (sourceId) result = result.sourceId(sourceId);
      return result;
    }).catch(err => console.error('Search error:', err));
  }, [search, selectedVertical]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="SearchAndApps" style={{ marginRight: '8px' }} />
        Example 2: Advanced Search with Builder
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
        {['All', 'Documents', 'Pages', 'People', 'Videos'].map((v) => (
          <DefaultButton
            key={v}
            text={v}
            onClick={() => setSelectedVertical(v)}
            primary={selectedVertical === v}
          />
        ))}
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 6 }}>
        <PrimaryButton
          text="Search All"
          onClick={() => handleSearch('*')}
          disabled={loading}
        />
        <DefaultButton
          text="Search Documents"
          onClick={() => handleSearch('filetype:docx OR filetype:pdf')}
          disabled={loading}
        />
      </Stack>

      {totalResults !== undefined && (
        <MessageBar messageBarType={MessageBarType.info}>
          Found {totalResults} result{totalResults !== 1 ? 's' : ''} in &quot;{selectedVertical}&quot;
        </MessageBar>
      )}

      {results.length > 0 && (
        <Stack tokens={{ childrenGap: 6 }}>
          {results.map((result) => (
            <Stack key={result.id} tokens={{ childrenGap: 2 }} styles={{ root: { padding: '10px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
              <div style={{ fontWeight: 600 }}>{result.data.Title || '(No Title)'}</div>
              <Label>{result.data.Author} - {result.data.LastModifiedTime ? new Date(result.data.LastModifiedTime).toLocaleDateString() : 'N/A'}</Label>
              <a href={result.data.Path} target="_blank" rel="noopener noreferrer" style={{ fontSize: '11px' }}>
                {result.data.Path}
              </a>
            </Stack>
          ))}
          {hasMore && (
            <PrimaryButton text={loading ? 'Loading...' : 'Load More'} onClick={loadMore} disabled={loading} />
          )}
        </Stack>
      )}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Builder API with verticals (People, Documents, Pages, Videos) and pagination.
      </Label>
    </Stack>
  );
};
