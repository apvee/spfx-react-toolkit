import * as React from 'react';
import {
  Stack,
  TextField,
  MessageBar,
  MessageBarType,
  Separator,
  Label,
  Icon,
} from '@fluentui/react';
import { useSPFxPnPSearch } from '../../../../hooks';

/**
 * Example 6: Search Suggestions (Autocomplete)
 */
export const PnPSearchSuggestionsDemo: React.FC = () => {
  const [query, setQuery] = React.useState('');
  const [suggestions, setSuggestions] = React.useState<string[]>([]);
  const { suggest } = useSPFxPnPSearch();
  
  // Track if component is mounted to prevent state updates after unmount
  const mountedRef = React.useRef(true);
  
  // Cleanup on unmount
  React.useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
    };
  }, []);

  const handleSuggest = React.useCallback(async (value: string) => {
    if (!value || value.length < 2) {
      if (mountedRef.current) {
        setSuggestions([]);
      }
      return;
    }
    try {
      const results = await suggest(value);
      if (mountedRef.current) {
        setSuggestions(results);
      }
    } catch (err) {
      console.error('[PnPSearchSuggestions] Error:', err);
      if (mountedRef.current) {
        setSuggestions([]);
      }
    }
  }, [suggest]);

  // Debounced suggest with proper cleanup
  const debouncedSuggest = React.useMemo(() => {
    let timeoutId: number | undefined;
    
    // Return debounced function
    const debounced = (value: string): void => {
      if (timeoutId !== undefined) {
        clearTimeout(timeoutId);
      }
      timeoutId = window.setTimeout(() => handleSuggest(value), 300);
    };
    
    // Cleanup function stored in the debounced function
    debounced.cancel = (): void => {
      if (timeoutId !== undefined) {
        clearTimeout(timeoutId);
        timeoutId = undefined;
      }
    };
    
    return debounced;
  }, [handleSuggest]);
  
  // Cleanup timeout on unmount
  React.useEffect(() => {
    return () => {
      debouncedSuggest.cancel?.();
    };
  }, [debouncedSuggest]);

  const handleQueryChange = React.useCallback((value: string) => {
    setQuery(value);
    debouncedSuggest(value);
  }, [debouncedSuggest]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="SearchBookmark" style={{ marginRight: '8px' }} />
        Example 4: Search Suggestions (Autocomplete)
      </h3>
      <Separator />

      <TextField
        label="Search with Autocomplete"
        value={query}
        onChange={(_, newValue) => handleQueryChange(newValue ?? '')}
        placeholder="Type at least 2 characters..."
      />

      {suggestions.length > 0 ? (
        <Stack tokens={{ childrenGap: 2 }} styles={{ root: { padding: '8px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
          <Label>Suggestions ({suggestions.length}):</Label>
          {suggestions.map((s, idx) => (
            <div
              key={idx}
              style={{ padding: '6px', backgroundColor: '#fff', borderRadius: '2px', cursor: 'pointer' }}
              onClick={() => setQuery(s)}
            >
              {s}
            </div>
          ))}
        </Stack>
      ) : query.length >= 2 ? (
        <MessageBar messageBarType={MessageBarType.info}>
          No suggestions available. This is normal in dev/test environments where there isn&apos;t enough indexed content.
          Try in a production environment with more search activity.
        </MessageBar>
      ) : null}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        Real-time autocomplete with 300ms debounce. Requires indexed content and search activity.
      </Label>
    </Stack>
  );
};
