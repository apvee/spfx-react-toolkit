import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Button, Field, Input, Label, MessageBar, MessageBarBody, Spinner } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPSearch } from '@apvee/spfx-react-toolkit';
type SuggestionState = {
    status: 'idle';
} | {
    status: 'loading';
} | {
    status: 'ready';
    suggestions: string[];
} | {
    status: 'error';
    message: string;
};
/** Example 4: Search Suggestions (Autocomplete) with distinct request states. */
export const PnPSearchSuggestionsDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const [query, setQuery] = React.useState('');
    const [suggestionState, setSuggestionState] = React.useState<SuggestionState>({ status: 'idle' });
    const { suggest } = useSPFxPnPSearch();
    const currentSuggest = React.useRef(suggest);
    currentSuggest.current = suggest;
    // Track if component is mounted to prevent state updates after unmount
    const mountedRef = React.useRef(true);
    const requestIdRef = React.useRef(0);
    // Cleanup on unmount
    React.useEffect(() => {
        mountedRef.current = true;
        return () => {
            mountedRef.current = false;
            requestIdRef.current += 1;
        };
    }, []);
    const handleSuggest = React.useCallback(async (value: string, requestId: number) => {
        if (value.length < 2 || !mountedRef.current || requestId !== requestIdRef.current) {
            return;
        }
        try {
            const results = await suggest(value);
            if (mountedRef.current && currentSuggest.current === suggest && requestId === requestIdRef.current) {
                setSuggestionState({ status: 'ready', suggestions: results });
            }
        }
        catch (err) {
            if (mountedRef.current && currentSuggest.current === suggest && requestId === requestIdRef.current) {
                console.error('[PnPSearchSuggestions] Error:', err);
                setSuggestionState({
                    status: 'error',
                    message: err instanceof Error && err.message ? err.message : 'The suggestion request failed.'
                });
            }
        }
    }, [suggest]);
    React.useEffect(() => { requestIdRef.current += 1; setSuggestionState({ status: 'idle' }); }, [suggest]);
    // Debounced suggest with proper cleanup
    const debouncedSuggest = React.useMemo(() => {
        let timeoutId: number | undefined;
        // Return debounced function
        const debounced = (value: string, requestId: number): void => {
            if (timeoutId !== undefined) {
                clearTimeout(timeoutId);
            }
            timeoutId = window.setTimeout(() => {
                timeoutId = undefined;
                return handleSuggest(value, requestId);
            }, 300);
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
        const requestId = ++requestIdRef.current;
        setQuery(value);
        if (value.length < 2) {
            debouncedSuggest.cancel();
            setSuggestionState({ status: 'idle' });
            return;
        }
        // The pending state includes debounce time; only a settled response can be empty.
        setSuggestionState({ status: 'loading' });
        debouncedSuggest(value, requestId);
    }, [debouncedSuggest]);
    return (<DemoCard title="Example 4: Search Suggestions (Autocomplete)" description="Type at least two characters to request SharePoint query suggestions." iconName="SearchBookmark" controls={<>{selector?.(false)}
    <Field label={"Search with Autocomplete"}>
      <Input value={query} onChange={(_, data) => handleQueryChange(data.value ?? '')} placeholder="Type at least 2 characters..."/>
    </Field>
  </>} contentPlacement="configuration" diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      Real-time autocomplete with 300ms debounce. Retrieves SharePoint query suggestions, not document title completions.
    </Label>
  </>}>

    {suggestionState.status === 'loading' && <Spinner size="tiny" label="Loading suggestions..."/>}

    {suggestionState.status === 'error' && <MessageBar intent="error">
      <MessageBarBody>
        Could not load suggestions. {suggestionState.message} Edit the search text to try again.
      </MessageBarBody>
    </MessageBar>}

    {suggestionState.status === 'ready' && (suggestionState.suggestions.length > 0 ? (<div className={demoStyles.panel}>
      <Label>Suggestions ({suggestionState.suggestions.length}):</Label>
      {suggestionState.suggestions.map((s, idx) => (<Button key={idx} onClick={() => {
                    requestIdRef.current += 1;
                    debouncedSuggest.cancel();
                    setQuery(s);
                    setSuggestionState({ status: 'idle' });
                }}>
        {s}
      </Button>))}
    </div>) : (<MessageBar intent="info">
      <MessageBarBody>
        No suggestions available for this text. SharePoint query suggestions depend on search history and tenant configuration.
        Indexed documents alone do not guarantee suggestions.
      </MessageBarBody>
    </MessageBar>))}

  </DemoCard>);
};
