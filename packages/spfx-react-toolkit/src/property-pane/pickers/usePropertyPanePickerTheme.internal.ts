import * as React from 'react';
import { useIsomorphicLayoutEffect } from '@fluentui/react-utilities';
import { webLightTheme } from '@fluentui/react-theme';
import type { Theme } from '@fluentui/react-theme';
import { useSPFxThemeInfo } from '../../hooks/useSPFxThemeInfo';
import { readPropertyPanePickerTheme, samePropertyPanePickerTheme } from './property-pane-picker-theme.internal';

/** Native picker-only bridge; SPFx theme notifications invalidate reads, never supply section colors. */
export function usePropertyPanePickerTheme(container: HTMLElement): Theme {
  useSPFxThemeInfo();
  const [snapshot, setSnapshot] = React.useState(() => ({ container, theme: webLightTheme }));
  const subscription = React.useRef<{ refresh: () => void }>();
  useIsomorphicLayoutEffect(() => {
    const document = container.ownerDocument, owner = document.defaultView;
    if (!owner) return;
    let disposed = false, frame: number | undefined;
    let ancestors: Element[] = [];
    const Observer = owner.MutationObserver;
    const schedule = (): void => {
      if (disposed || frame !== undefined) return;
      if (typeof owner.requestAnimationFrame !== 'function') { refresh(); return; }
      frame = owner.requestAnimationFrame(() => { frame = undefined; refresh(); });
    };
    const pathObserver = typeof Observer === 'function' ? new Observer(schedule) : undefined;
    const connectionObserver = typeof Observer === 'function' ? new Observer(schedule) : undefined;
    const isStylesheet = (node: Node): boolean => node.nodeName === 'STYLE' || node.nodeName === 'LINK' || node.parentNode?.nodeName === 'STYLE';
    const styleObserver = typeof Observer === 'function' ? new Observer(records => {
      // Native Fluent replaces theme style elements. CSSOM-only edits require a
      // host theme notification or native onRender; there is no polling/patching.
      if (records.some(record => isStylesheet(record.target) || [...Array.from(record.addedNodes), ...Array.from(record.removedNodes)].some(isStylesheet))) schedule();
    }) : undefined;
    function bindAncestors(): void {
      if (!container.isConnected) {
        pathObserver?.disconnect(); ancestors = [];
        // This subtree observation exists only until SPFx attaches its detached field.
        connectionObserver?.observe(document.documentElement, { childList: true, subtree: true });
        return;
      }
      connectionObserver?.disconnect();
      const next: Element[] = [];
      for (let node: Element | undefined = container; node; node = node.parentElement ?? undefined) next.push(node);
      if (next.length === ancestors.length && next.every((node, index) => node === ancestors[index])) return;
      ancestors = next; pathObserver?.disconnect();
      for (const node of ancestors) pathObserver?.observe(node, {
        attributes: true, attributeFilter: ['class', 'style', 'dir'], childList: node !== container,
      });
    }
    function refresh(): void {
      if (disposed) return;
      bindAncestors();
      const next = readPropertyPanePickerTheme(container);
      setSnapshot(previous => {
        const sameHost = previous.container === container;
        const theme = next ?? (sameHost ? previous.theme : webLightTheme);
        return sameHost && samePropertyPanePickerTheme(previous.theme, theme) ? previous : { container, theme };
      });
    }
    const current = { refresh }; subscription.current = current;
    bindAncestors();
    const onStylesheetLoad = (event: Event): void => {
      if (event.target instanceof owner.Element && event.target.nodeName === 'LINK') schedule();
    };
    if (document.head) {
      styleObserver?.observe(document.head, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['href', 'media', 'disabled'] });
      document.head.addEventListener('load', onStylesheetLoad, true);
    }
    owner.addEventListener('resize', schedule);
    return () => {
      disposed = true;
      if (frame !== undefined) owner.cancelAnimationFrame(frame);
      pathObserver?.disconnect(); connectionObserver?.disconnect(); styleObserver?.disconnect();
      document.head?.removeEventListener('load', onStylesheetLoad, true);
      owner.removeEventListener('resize', schedule);
      ancestors = [];
      if (subscription.current === current) subscription.current = undefined;
    };
  }, [container]);
  // Includes native onRender updates and SPFx invalidations. Stable snapshots
  // bail out; our own provider stylesheet can cause only a bounded re-read.
  useIsomorphicLayoutEffect(() => { subscription.current?.refresh(); });
  return snapshot.container === container ? snapshot.theme : webLightTheme;
}
