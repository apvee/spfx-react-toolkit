import * as React from 'react';
import { FluentProvider } from '@fluentui/react-components';
import { ThemeContext_unstable, useFluent_unstable, useThemeClassName_unstable } from '@fluentui/react-shared-contexts';
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

function ShellLayout(props: RuntimeProps & { resolvedSurface: 'section' | 'application' }): React.ReactElement {
  const { ownerDocument, rootRef, resolvedSurface, padding = 'auto',
    action, children, className, style, ...nativeProps } = props;
  delete nativeProps.dangerouslySetInnerHTML;
  delete nativeProps.surface;
  delete nativeProps.theme;
  delete nativeProps.dir;
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  const { isEdit } = useSPFxDisplayMode();
  const styles = useWebPartShellLayoutStyles(resolvedSurface, padding);
  const changeTitle = React.useCallback((update: { title: string; titleHTML: string }): void => {
    setProperties(update);
  }, [setProperties]);
  return <div {...nativeProps} ref={rootRef} className={[styles.root, className].filter(Boolean).join(' ')} style={style}
    data-web-part-shell-surface={resolvedSurface}>
    <div className={styles.content} data-web-part-shell-content="">
      <div className={styles.header}>
        <div className={styles.title}>
          <WebPartTitle title={properties?.title ?? ''} titleHTML={properties?.titleHTML}
            showTitle={properties?.showTitle ?? true} editing={isEdit} onChange={changeTitle} ownerDocument={ownerDocument} />
        </div>
        <div className={styles.actions}>{action}</div>
      </div>
      {children}
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
  return <FluentProvider theme={theme} dir={dir} targetDocument={props.ownerDocument} className={styles.provider}>
    <ShellLayout {...props} resolvedSurface={surface} />
  </FluentProvider>;
}
