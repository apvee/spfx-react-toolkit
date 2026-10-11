import type { TeamsSdkLike } from '../helpers/spfx-host-context.internal';
import type { SPFxRuntimeStore, SPFxRuntimeTeamsTheme } from './runtime-store.internal';

interface StoreSubscription {
  store: SPFxRuntimeStore | undefined;
  references: number;
  liveTheme?: SPFxRuntimeTeamsTheme;
}

interface TeamsRuntimeBroker {
  readonly stores: Map<SPFxRuntimeStore, StoreSubscription>;
}

// Cache allocation has no SDK effects. Registration is lazy and scoped to this
// module's SDK identities; independently bundled SDK consumers can replace it.
const brokers = new WeakMap<TeamsSdkLike, TeamsRuntimeBroker>();

function normalizeTheme(value: unknown): SPFxRuntimeTeamsTheme {
  const theme = typeof value === 'string' ? value.toLowerCase() : undefined;
  return theme === 'dark' ? 'dark'
    : theme === 'contrast' || theme === 'highcontrast' ? 'highContrast'
    : 'default';
}

function readInitialTheme(context: unknown, v2: boolean): unknown {
  if (typeof context !== 'object' || context === null) {
    return undefined;
  }
  const record = context as Record<string, unknown>;
  if (!v2) {
    return record.theme;
  }
  const app = record.app;
  return typeof app === 'object' && app !== null
    ? (app as Record<string, unknown>).theme
    : undefined;
}

function registerThemeDispatcher(sdk: TeamsSdkLike, broker: TeamsRuntimeBroker): void {
  const dispatch = (value: string): void => {
    const theme = normalizeTheme(value);
    for (const subscription of broker.stores.values()) {
      const store = subscription.store;
      if (store) {
        subscription.liveTheme = theme;
        store.setState({ teams: { ...store.getState().teams, theme } });
      }
    }
  };

  if (typeof sdk.app?.registerOnThemeChangeHandler === 'function') {
    try {
      sdk.app.registerOnThemeChangeHandler(dispatch);
      return;
    } catch (error) {
      console.error('[useSPFxTeams] Unable to register v2 theme handler.', error);
    }
  }
  if (typeof sdk.registerOnThemeChangeHandler === 'function') {
    try {
      sdk.registerOnThemeChangeHandler(dispatch);
    } catch (error) {
      console.error('[useSPFxTeams] Unable to register v1 theme handler.', error);
    }
  }
}

async function initializeStore(sdk: TeamsSdkLike, subscription: StoreSubscription): Promise<void> {
  const apply = (context: unknown, v2: boolean): void => {
    subscription.store?.setState({
      teams: {
        supported: true,
        initialized: true,
        context,
        theme: subscription.liveTheme ?? normalizeTheme(readInitialTheme(context, v2)),
      },
    });
  };

  if (typeof sdk.app?.getContext === 'function') {
    try {
      const context = await sdk.app.getContext();
      apply(context, true);
      return;
    } catch (error) {
      console.error('[useSPFxTeams] Unable to acquire v2 context.', error);
    }
  }
  if (!subscription.store) {
    return;
  }
  if (typeof sdk.getContext === 'function') {
    try {
      sdk.getContext(context => apply(context, false));
      return;
    } catch (error) {
      console.error('[useSPFxTeams] Unable to acquire v1 context.', error);
    }
  }
  subscription.store?.setState({ teams: { supported: false, initialized: true } });
}

/**
 * Share one SDK dispatcher and one acquisition per active runtime store.
 * The SDK exposes no unregister method: cleanup removes stores and leaves an
 * inert dispatcher. A subsequent activation reuses the registered dispatcher.
 */
export function subscribeSPFxTeamsRuntime(sdk: TeamsSdkLike, store: SPFxRuntimeStore): () => void {
  const existingBroker = brokers.get(sdk);
  const broker = existingBroker ?? { stores: new Map<SPFxRuntimeStore, StoreSubscription>() };
  if (!existingBroker) {
    brokers.set(sdk, broker);
  }
  const existingSubscription = broker.stores.get(store);
  const subscription = existingSubscription ?? { store, references: 1 };
  if (existingSubscription) {
    subscription.references++;
  } else {
    broker.stores.set(store, subscription);
    store.setState({ teams: { supported: false, initialized: false } });
    if (!existingBroker) {
      registerThemeDispatcher(sdk, broker);
    }
    initializeStore(sdk, subscription).catch(error => {
      console.error('[useSPFxTeams] Unexpected context initialization failure.', error);
    });
  }

  let subscribed = true;
  return () => {
    if (!subscribed) {
      return;
    }
    subscribed = false;
    subscription.references--;
    if (subscription.references === 0) {
      broker.stores.delete(store);
      subscription.store = undefined;
    }
  };
}
