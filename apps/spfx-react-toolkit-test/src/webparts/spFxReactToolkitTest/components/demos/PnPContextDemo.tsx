import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';
import { DemoIcon } from '../shared/DemoIcon';
import { Button, Field, Input, Label, MessageBar, MessageBarActions, MessageBarBody } from '@fluentui/react-components';
import * as React from 'react';
import { useSPFxPnPContext } from '@apvee/spfx-react-toolkit';
import { InfoRow } from '../shared';
import '@pnp/sp/webs';
/**
 * Example 1: useSPFxPnPContext
 * Shows how to create PnP contexts for current and cross-site scenarios
 */
export const PnPContextDemo: React.FC<{
    readonly selector?: (busy?: boolean) => React.ReactNode;
}> = ({ selector }) => {
    const demoStyles = useDemoStyles();
    const [crossSiteUrl, setCrossSiteUrl] = React.useState('');
    const [siteInfo, setSiteInfo] = React.useState<{
        title: string;
        url: string;
        description: string;
    } | undefined>(undefined);
    const [loading, setLoading] = React.useState(false);
    const [errorMsg, setErrorMsg] = React.useState<string | undefined>(undefined);
    // Current site context (default)
    const currentContext = useSPFxPnPContext();
    // Cross-site context (only when URL is provided)
    const crossSiteContext = useSPFxPnPContext(crossSiteUrl || undefined);
    const context = crossSiteUrl ? crossSiteContext : currentContext;
    const identity = context.sp;
    const target = crossSiteUrl;
    const current = React.useRef({ identity, target });
    current.current = { identity, target };
    const mounted = React.useRef(true);
    const request = React.useRef(0);
    React.useEffect(() => { mounted.current = true; return () => { mounted.current = false; }; }, []);
    React.useEffect(() => { setSiteInfo(undefined); setErrorMsg(undefined); setLoading(false); }, [identity, target]);
    const handleLoadCurrentSite = React.useCallback(async () => {
        if (!identity)
            return;
        const id = ++request.current;
        const active = (): boolean => mounted.current && request.current === id && current.current.identity === identity && current.current.target === target;
        setLoading(true);
        setSiteInfo(undefined);
        setErrorMsg(undefined);
        try {
            const web = await identity.web.select('Title', 'Url', 'Description')();
            if (active())
                setSiteInfo({ title: web.Title, url: web.Url, description: web.Description || '(no description)' });
        }
        catch (error) {
            console.error('Error loading site:', error);
            if (active())
                setErrorMsg(error instanceof Error ? error.message : 'Unknown error');
        }
        finally {
            if (active())
                setLoading(false);
        }
    }, [identity, target]);
    return (<DemoCard title="Example 1: useSPFxPnPContext - Site Information" description="Load site information using the current or a cross-site PnP context." iconName="Globe" controls={<>{selector?.(false)}

    <Field label={"Site URL"} hint={"Leave empty to use current site"}>
      <Input value={crossSiteUrl} onChange={(_, data) => setCrossSiteUrl(data.value ?? '')} placeholder="e.g., /sites/hr or https://contoso.sharepoint.com/sites/hr"/>
    </Field>
  </>} contentPlacement="configuration" commands={<>
    <Button appearance="primary" onClick={handleLoadCurrentSite} disabled={!context.isInitialized || loading}>{loading ? 'Loading...' : 'Run'}</Button>

  </>} diagnostics={<>
    <Label>
      <DemoIcon name={"InfoSolid"}/>
      useSPFxPnPContext creates configured SPFI instances. Use for cross-site scenarios or when you need custom cache/batch config.
    </Label>
  </>}>

    {errorMsg && (<MessageBar intent="error">
      <MessageBarBody>
        {errorMsg}
      </MessageBarBody>
      <MessageBarActions>
        <Button aria-label="Dismiss message" onClick={() => setErrorMsg(undefined)}>Dismiss</Button>
      </MessageBarActions>
    </MessageBar>)}

    {!siteInfo && !errorMsg && <Label>{loading ? 'Loading...' : 'Run to read the selected site.'}</Label>}
    {/* Results */}

    {siteInfo && (<div className={demoStyles.panel}>
      <Label className={demoStyles.observations}>
        <DemoIcon name={"CheckMark"}/>
        Site Information Loaded:
      </Label>
      <InfoRow label="Title" value={siteInfo.title} icon="CityNext"/>
      <InfoRow label="URL" value={siteInfo.url} icon="Link"/>
      <InfoRow label="Description" value={siteInfo.description} icon="Info"/>
    </div>)}

  </DemoCard>);
};
