import * as React from 'react';
import { Card, FluentProvider } from '@fluentui/react-components';
import { ThemeContext_unstable, useFluent_unstable, useThemeClassName_unstable } from '@fluentui/react-shared-contexts';
import { renderContentSlot } from '../content-slot.internal';
import { useSx } from '../../hooks/useSx';
import { useSPFxProperties } from '../../hooks/useSPFxProperties';
import { useSPFxDisplayMode } from '../../hooks/useSPFxDisplayMode';
import { useSPFxContext } from '../../core/context.internal';
import { useSPFxFluent9ThemeInfo } from '../../hooks/useSPFxFluent9ThemeInfo';
import { isSingleWebPartAppPage, resolveSPFxTeamsSdk, resolveWebPartShellSurface } from '../../helpers/spfx-host-context.internal';
import { WebPartTitle } from './title/web-part-title.internal';
import { useWebPartShellLayoutStyles, useWebPartShellProviderStyles } from './web-part-shell.styles.internal';
import type { IWebPartShellProperties, WebPartShellProps } from './web-part-shell.types';

type RuntimeProps = Omit<WebPartShellProps<IWebPartShellProperties>, 'instance'> & {
  ownerDocument: Document;
  rootRef: React.ForwardedRef<HTMLDivElement>;
};

function ShellLayout(props: Omit<RuntimeProps, 'ownerDocument'> & { resolvedSurface: 'section' | 'application' }): React.ReactElement {
  const { rootRef, resolvedSurface, strings, titleEditorStrings, padding = 'auto',
    appearance = 'subtle', size = 'medium', titlePosition = 'outside', action, children, className, style, ...nativeProps } = props;
  delete nativeProps.dangerouslySetInnerHTML;
  delete nativeProps.surface;
  delete nativeProps.theme;
  delete nativeProps.dir;
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  const { isEdit } = useSPFxDisplayMode();
  const styles = useWebPartShellLayoutStyles(resolvedSurface, padding);
  const sx = useSx();
  const showTitle = properties?.showTitle ?? true;
  const changeTitle = React.useCallback((update: { title: string; titleHTML: string }): void => {
    setProperties(update);
  }, [setProperties]);
  const header = showTitle && <div key="header" className={styles.header} data-web-part-shell-header="">
    <div className={styles.title}>
      <WebPartTitle title={properties?.title ?? ''} titleHTML={properties?.titleHTML}
        showTitle={showTitle} editing={isEdit} onChange={changeTitle} strings={strings} titleEditorStrings={titleEditorStrings} />
    </div>
    <div className={styles.actions}>{renderContentSlot(action, 'div', sx, { shorthand: 'preserve' })}</div>
  </div>;
  const body = <div key="body" className={styles.body} data-web-part-shell-body="">{children}</div>;
  const hasCard = appearance !== 'subtle';
  return <div {...nativeProps} ref={rootRef} className={[styles.root, className].filter(Boolean).join(' ')} style={style}
    data-web-part-shell-surface={resolvedSurface}>
    <div className={styles.content} data-web-part-shell-content="">
      {(!hasCard || titlePosition === 'outside') && header}
      {hasCard ? <Card key="card" appearance={appearance} size={size} orientation="vertical" focusMode="off"
        data-web-part-shell-card="">
        {titlePosition === 'inside' && header}
        {body}
      </Card> : body}
    </div>
  </div>;
}

/** Host theme conversion/subscription and SDK detection remain owned by existing hooks/helpers. */
export function WebPartShellRuntime(props: RuntimeProps): React.ReactElement {
  const { spfxContext } = useSPFxContext();
  const { theme: hostTheme } = useSPFxFluent9ThemeInfo();
  const ancestorTheme = React.useContext(ThemeContext_unstable);
  const ancestorFluent = useFluent_unstable();
  // The public theme class context distinguishes a FluentProvider from useFluent's LTR default.
  const hasAncestorProvider = useThemeClassName_unstable() !== '';
  const theme = React.useMemo(() => ({ ...hostTheme, ...ancestorTheme, ...props.theme }), [hostTheme, ancestorTheme, props.theme]);
  const dir = props.dir ?? (hasAncestorProvider ? ancestorFluent.dir
    : spfxContext.pageContext.cultureInfo.isRightToLeft ? 'rtl' : 'ltr');
  const surface = resolveWebPartShellSurface(props.surface ?? 'auto', resolveSPFxTeamsSdk(spfxContext) !== undefined,
    isSingleWebPartAppPage(spfxContext.pageContext));
  const styles = useWebPartShellProviderStyles(dir);
  const { ownerDocument, ...layoutProps } = props;
  return <FluentProvider theme={theme} dir={dir} targetDocument={ownerDocument} className={styles.provider}>
    <ShellLayout {...layoutProps} resolvedSurface={surface} />
  </FluentProvider>;
}
