import * as React from 'react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { DisplayMode } from '@microsoft/sp-core-library';
import type { ContainerSize } from './types';
import type {
  SPFxRuntimeState,
  SPFxRuntimeStore,
  SPFxRuntimeTeamsState,
} from './runtime-store.internal';

export const SPFxRuntimeStoreContext = React.createContext<SPFxRuntimeStore | undefined>(undefined);

if (process.env.NODE_ENV !== 'production') {
  SPFxRuntimeStoreContext.displayName = 'SPFxRuntimeStoreContext';
}

function objectIs<T>(previous: T, next: T): boolean {
  return Object.is(previous, next);
}

export function useSPFxRuntimeStore(): SPFxRuntimeStore {
  const store = React.useContext(SPFxRuntimeStoreContext);

  if (!store) {
    throw new Error(
      'SPFx runtime state is not available. ' +
      'Make sure your component is wrapped with a host-specific SPFx provider.'
    );
  }

  return store;
}

export function useSPFxRuntimeSelector<T>(
  selector: (state: SPFxRuntimeState) => T,
  isEqual: (previous: T, next: T) => boolean = objectIs
): T {
  const store = useSPFxRuntimeStore();
  const selectorRef = React.useRef(selector);
  const equalityRef = React.useRef(isEqual);
  const [selected, setSelected] = React.useState<T>(() => selector(store.getState()));

  React.useEffect(() => {
    selectorRef.current = selector;
    equalityRef.current = isEqual;

    setSelected((previous): T => {
      const next = selector(store.getState());
      return isEqual(previous, next) ? previous : next;
    });
  }, [selector, isEqual, store]);

  React.useEffect(() => {
    const checkForUpdates = (): void => {
      setSelected((previous): T => {
        const next = selectorRef.current(store.getState());
        return equalityRef.current(previous, next) ? previous : next;
      });
    };

    const unsubscribe = store.subscribe(checkForUpdates);
    checkForUpdates();

    return unsubscribe;
  }, [store]);

  return selected;
}

export interface SPFxRuntimeActions {
  readonly setTheme: (theme: IReadonlyTheme | undefined) => void;
  readonly setDisplayMode: (displayMode: DisplayMode | undefined) => void;
  readonly setProperties: (updater: unknown | ((previous: unknown) => unknown)) => void;
  readonly setContainerElement: (element: HTMLElement | undefined) => void;
  readonly setContainerSize: (size: ContainerSize | undefined) => void;
  readonly setTeamsState: (teams: SPFxRuntimeTeamsState) => void;
}

export function createSPFxRuntimeActions(store: SPFxRuntimeStore): SPFxRuntimeActions {
  return {
    setTheme(theme: IReadonlyTheme | undefined): void {
      store.setState({ theme });
    },
    setDisplayMode(displayMode: DisplayMode | undefined): void {
      store.setState({ displayMode });
    },
    setProperties(updater: unknown | ((previous: unknown) => unknown)): void {
      store.setState((previous): SPFxRuntimeState => {
        const properties = typeof updater === 'function'
          ? (updater as (previous: unknown) => unknown)(previous.properties)
          : updater;

        if (Object.is(properties, previous.properties)) {
          return previous;
        }

        return {
          ...previous,
          properties,
        };
      });
    },
    setContainerElement(element: HTMLElement | undefined): void {
      store.setState({ containerEl: element });
    },
    setContainerSize(size: ContainerSize | undefined): void {
      store.setState({ containerSize: size });
    },
    setTeamsState(teams: SPFxRuntimeTeamsState): void {
      store.setState({ teams });
    },
  };
}

export function useSPFxRuntimeActions(): SPFxRuntimeActions {
  const store = useSPFxRuntimeStore();
  return React.useMemo(() => createSPFxRuntimeActions(store), [store]);
}
