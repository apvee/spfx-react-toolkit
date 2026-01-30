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
import { useSPFxPnPSearch } from '../../../../hooks';

/**
 * Example 5: Faceted Search with Refiners
 */
export const PnPSearchRefinersDemo: React.FC = () => {
  const {
    search,
    results,
    refiners,
    totalResults,
    loading,
    applyRefiner,
    error,
    clearError,
  } = useSPFxPnPSearch<{ Title: string; Path: string; FileType: string; Author: string }>({
    pageSize: 10,
    selectProperties: ['Title', 'Path', 'FileType', 'Author'],
    refiners: 'FileType,Author'
  });

  const handleSearch = React.useCallback(() => {
    search('*').catch(err => console.error('Search error:', err));
  }, [search]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="FilterSettings" style={{ marginRight: '8px' }} />
        Example 3: Faceted Search with Refiners
      </h3>
      <Separator />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={clearError}>
          {error.message}
        </MessageBar>
      )}

      <PrimaryButton
        text={loading ? 'Loading...' : 'Search All Content'}
        onClick={handleSearch}
        disabled={loading}
        iconProps={{ iconName: 'Search' }}
      />

      {totalResults !== undefined && (
        <Label>
          <Icon iconName="SearchIssue" style={{ marginRight: '4px', color: '#0078d4' }} />
          Found {totalResults} results
        </Label>
      )}

      <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { alignItems: 'flex-start' } }}>
        {refiners.length > 0 && (
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: '200px', padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
            <Label style={{ fontWeight: 600 }}>Filters</Label>
            {refiners.map((refiner) => (
              <Stack key={refiner.name} tokens={{ childrenGap: 2 }}>
                <Label>{refiner.name}</Label>
                {refiner.entries.slice(0, 5).map((entry) => (
                  <DefaultButton
                    key={entry.value}
                    text={`${entry.value} (${entry.count})`}
                    onClick={() => applyRefiner(refiner.name, entry.value)}
                    styles={{ root: { height: '28px', fontSize: '12px' } }}
                  />
                ))}
              </Stack>
            ))}
          </Stack>
        )}

        <Stack tokens={{ childrenGap: 6 }} styles={{ root: { flexGrow: 1 } }}>
          {results.length > 0 ? (
            results.map((result) => (
              <Stack key={result.id} tokens={{ childrenGap: 2 }} styles={{ root: { padding: '8px', backgroundColor: '#faf9f8', borderRadius: '4px' } }}>
                <div style={{ fontWeight: 600 }}>{result.data.Title || '(No Title)'}</div>
                <Label>{result.data.FileType} | {result.data.Author}</Label>
                <a href={result.data.Path} target="_blank" rel="noopener noreferrer" style={{ fontSize: '11px' }}>
                  {result.data.Path}
                </a>
              </Stack>
            ))
          ) : (
            <Label>Click &quot;Search All Content&quot; to see results.</Label>
          )}
        </Stack>
      </Stack>

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Faceted search with FileType and Author refiners. Click to filter/toggle.
      </Label>
    </Stack>
  );
};
