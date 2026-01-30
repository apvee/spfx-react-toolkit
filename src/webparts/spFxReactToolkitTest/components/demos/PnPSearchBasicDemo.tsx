import * as React from 'react';
import {
  Stack,
  TextField,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';
import { useSPFxPnPSearch } from '../../../../hooks';

/**
 * Example 3: useSPFxPnPSearch - Basic Search
 */
export const PnPSearchBasicDemo: React.FC = () => {
  const [searchText, setSearchText] = React.useState('');
  
  const {
    search,
    results,
    totalResults,
    loading,
    error,
    clearError,
  } = useSPFxPnPSearch<{ Title: string; Path: string; FileType: string }>({
    pageSize: 10
  });

  const handleSearch = React.useCallback(async () => {
    if (!searchText.trim()) return;
    try {
      await search(searchText);
    } catch (err) {
      console.error('Search error:', err);
    }
  }, [searchText, search]);


  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="Search" style={{ marginRight: '8px' }} />
        Example 1: Basic Search
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <TextField
          value={searchText}
          onChange={(_, newValue) => setSearchText(newValue ?? '')}
          placeholder="Search SharePoint..."
          styles={{ root: { flexGrow: 1 } }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              handleSearch().catch(err => console.error('Search error:', err));
            }
          }}
        />
        <PrimaryButton
          text={loading ? 'Searching...' : 'Search'}
          onClick={handleSearch}
          disabled={!searchText.trim() || loading}
          iconProps={{ iconName: 'Search' }}
        />
      </Stack>

      {totalResults !== undefined && (
        <MessageBar messageBarType={MessageBarType.info}>
          Found {totalResults} result{totalResults !== 1 ? 's' : ''}
        </MessageBar>
      )}

      {results.length > 0 ? (
        <Stack tokens={{ childrenGap: 6 }}>
          {results.map((result) => (
            <Stack
              key={result.id}
              tokens={{ childrenGap: 4 }}
              styles={{ root: { padding: '10px', backgroundColor: '#faf9f8', borderRadius: '4px' } }}
            >
              <div style={{ fontWeight: 600 }}>{result.data.Title || '(No Title)'}</div>
              <a href={result.data.Path} target="_blank" rel="noopener noreferrer" style={{ fontSize: '12px' }}>
                {result.data.Path}
              </a>
              {result.data.FileType && <Label>Type: {result.data.FileType}</Label>}
            </Stack>
          ))}
        </Stack>
      ) : totalResults === 0 ? (
        <MessageBar>No results found. Try a different query.</MessageBar>
      ) : null}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Simple text search. Try: &quot;document&quot;, &quot;report&quot;, &quot;ContentType:Document&quot;
      </Label>
    </Stack>
  );
};
