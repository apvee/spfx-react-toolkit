import * as React from 'react';
import { FluentProvider } from '@fluentui/react-components';
import { ThemeContext_unstable, useFluent_unstable, useThemeClassName_unstable } from '@fluentui/react-shared-contexts';
import { hasContentSlot, inspectContentSlot, renderContentSlot } from '../content-slot.internal';
import { useSx } from '../../hooks/useSx';
import { useSPFxProperties } from '../../hooks/useSPFxProperties';
import { useSPFxDisplayMode } from '../../hooks/useSPFxDisplayMode';
import { useSPFxContext } from '../../core/context.internal';
import { useSPFxFluent9ThemeInfo } from '../../hooks/useSPFxFluent9ThemeInfo';
import { isSingleWebPartAppPage, resolveSPFxTeamsSdk, resolveWebPartShellSurface } from '../../helpers/spfx-host-context.internal';
import { resolveWebPartShellStrings } from './localization/web-part-shell-locales.internal';
import type { WebPartShellStrings } from './localization/web-part-shell-strings.internal';
import { WebPartTitle } from './title/web-part-title.internal';
import { useWebPartShellLayoutStyles, useWebPartShellProviderStyles } from './web-part-shell.styles.internal';
import type { IWebPartShellProperties, WebPartShellProps } from './web-part-shell.types';

type RuntimeProps = Omit<WebPartShellProps<IWebPartShellProperties>, 'instance'> & {
  ownerDocument: Document;
  rootRef: React.ForwardedRef<HTMLDivElement>;
};

function ShellLayout(props: RuntimeProps & { resolvedSurface: 'section' | 'application'; strings: WebPartShellStrings }): React.ReactElement {
  const { ownerDocument, rootRef, resolvedSurface, strings, padding = 'auto',
    action, children, className, style, ...nativeProps } = props;
  delete nativeProps.dangerouslySetInnerHTML;
  delete nativeProps.surface;
  delete nativeProps.theme;
  delete nativeProps.dir;
  const { properties, setProperties } = useSPFxProperties<IWebPartShellProperties>();
  const { isEdit } = useSPFxDisplayMode();
  const styles = useWebPartShellLayoutStyles(resolvedSurface, padding);
  const sx = useSx();
  const showTitle = properties?.showTitle ?? true;
  const actionInspection = inspectContentSlot(action);
  const hasHeader = showTitle || (actionInspection.kind === 'shorthand'
    ? React.Children.toArray(actionInspection.children).length > 0 : hasContentSlot(action));
  const changeTitle = React.useCallback((update: { title: string; titleHTML: string }): void => {
    setProperties(update);
  }, [setProperties]);
  return <div {...nativeProps} ref={rootRef} className={[styles.root, className].filter(Boolean).join(' ')} style={style}
    data-web-part-shell-surface={resolvedSurface}>
    <div className={styles.content} data-web-part-shell-content="">
      {hasHeader && <div className={styles.header}>
        <div className={styles.title}>
          <WebPartTitle title={properties?.title ?? ''} titleHTML={properties?.titleHTML}
            showTitle={showTitle} editing={isEdit} onChange={changeTitle} ownerDocument={ownerDocument} strings={strings} />
        </div>
        <div className={styles.actions}>{renderContentSlot(action, 'div', sx, { shorthand: 'preserve' })}</div>
      </div>}
      <div className={styles.body} data-web-part-shell-body="">{children}</div>
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
    <ShellLayout {...props} resolvedSurface={surface} strings={resolveWebPartShellStrings(spfxContext.pageContext.cultureInfo.currentUICultureName)} />
  </FluentProvider>;
}
