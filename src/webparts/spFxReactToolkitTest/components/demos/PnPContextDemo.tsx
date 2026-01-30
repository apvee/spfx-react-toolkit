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
import { useSPFxPnPContext } from '../../../../hooks';
import { InfoRow } from '../shared';
import '@pnp/sp/webs';

/**
 * Example 1: useSPFxPnPContext
 * Shows how to create PnP contexts for current and cross-site scenarios
 */
export const PnPContextDemo: React.FC = () => {
  const [crossSiteUrl, setCrossSiteUrl] = React.useState('');
  const [siteInfo, setSiteInfo] = React.useState<{ title: string; url: string; description: string } | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | null>(null);

  // Current site context (default)
  const currentContext = useSPFxPnPContext();

  // Cross-site context (only when URL is provided)
  const crossSiteContext = useSPFxPnPContext(crossSiteUrl || undefined);

  const handleLoadCurrentSite = React.useCallback(async () => {
    if (!currentContext.sp) return;

    setLoading(true);
    setErrorMsg(null);
    try {
      const web = await currentContext.sp.web.select('Title', 'Url', 'Description')();
      setSiteInfo({
        title: web.Title,
        url: web.Url,
        description: web.Description || '(no description)'
      });
    } catch (error) {
      console.error('Error loading site:', error);
      setErrorMsg(error instanceof Error ? error.message : 'Unknown error');
      setSiteInfo(null);
    } finally {
      setLoading(false);
    }
  }, [currentContext.sp]);

  const handleLoadCrossSite = React.useCallback(async () => {
    if (!crossSiteContext.sp || !crossSiteUrl) return;

    setLoading(true);
    setErrorMsg(null);
    try {
      const web = await crossSiteContext.sp.web.select('Title', 'Url', 'Description')();
      setSiteInfo({
        title: web.Title,
        url: web.Url,
        description: web.Description || '(no description)'
      });
    } catch (error) {
      console.error('Error loading cross-site:', error);
      setErrorMsg(error instanceof Error ? error.message : 'Unknown error');
      setSiteInfo(null);
    } finally {
      setLoading(false);
    }
  }, [crossSiteContext.sp, crossSiteUrl]);

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px', border: '1px solid #edebe9', borderRadius: '4px' } }}>
      <h3>
        <Icon iconName="Globe" style={{ marginRight: '8px' }} />
        Example 1: useSPFxPnPContext - Site Information
      </h3>
      <Separator />

      {errorMsg && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg(null)}>
          {errorMsg}
        </MessageBar>
      )}

      {/* Current Site Section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Current Site Context</Label>
        <InfoRow label="Effective URL" value={currentContext.siteUrl} icon="Link" />
        <InfoRow label="Is Initialized" value={currentContext.isInitialized ? 'Yes' : 'No'} icon="CheckMark" />
        {currentContext.error && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Init Error: {currentContext.error.message}
          </MessageBar>
        )}
        <PrimaryButton
          text={loading ? 'Loading...' : 'Load Current Site Info'}
          onClick={handleLoadCurrentSite}
          disabled={!currentContext.isInitialized || loading}
          iconProps={{ iconName: 'CloudDownload' }}
        />
      </Stack>

      {/* Cross-Site Section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Cross-Site Context (Optional)</Label>
        <TextField
          label="Site URL"
          value={crossSiteUrl}
          onChange={(_, newValue) => setCrossSiteUrl(newValue ?? '')}
          placeholder="e.g., /sites/hr or https://contoso.sharepoint.com/sites/hr"
          description="Leave empty to use current site"
        />
        {crossSiteUrl && (
          <>
            <InfoRow label="Resolved URL" value={crossSiteContext.siteUrl} icon="Link" />
            <InfoRow label="Is Initialized" value={crossSiteContext.isInitialized ? 'Yes' : 'No'} icon="CheckMark" />
            {crossSiteContext.error && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Init Error: {crossSiteContext.error.message}
              </MessageBar>
            )}
            <PrimaryButton
              text={loading ? 'Loading...' : 'Load Cross-Site Info'}
              onClick={handleLoadCrossSite}
              disabled={!crossSiteContext.isInitialized || loading}
              iconProps={{ iconName: 'Globe' }}
            />
          </>
        )}
      </Stack>

      {/* Results */}
      {siteInfo && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { padding: '12px', backgroundColor: '#f3f2f1', borderRadius: '4px' } }}>
          <Label style={{ fontWeight: 600 }}>
            <Icon iconName="CheckMark" style={{ marginRight: '4px', color: '#107c10' }} />
            Site Information Loaded:
          </Label>
          <InfoRow label="Title" value={siteInfo.title} icon="CityNext" />
          <InfoRow label="URL" value={siteInfo.url} icon="Link" />
          <InfoRow label="Description" value={siteInfo.description} icon="Info" />
        </Stack>
      )}

      <Label>
        <Icon iconName="InfoSolid" style={{ marginRight: '4px', color: '#0078d4' }} />
        useSPFxPnPContext creates configured SPFI instances. Use for cross-site scenarios or when you need custom cache/batch config.
      </Label>
    </Stack>
  );
};
