import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { DisplayMode } from '@microsoft/sp-core-library';
import type { ContainerSize } from './types';

export type SPFxRuntimeTeamsTheme = 'default' | 'dark' | 'highContrast';

export interface SPFxRuntimeTeamsState {
  readonly supported: boolean;
  readonly context?: unknown;
  readonly theme?: SPFxRuntimeTeamsTheme;
  readonly initialized: boolean;
}

export interface SPFxRuntimeState {
  readonly theme: IReadonlyTheme | undefined;
  readonly displayMode: DisplayMode | undefined;
  readonly properties: unknown;
  readonly containerEl: HTMLElement | undefined;
  readonly containerSize: ContainerSize | undefined;
  readonly teams: SPFxRuntimeTeamsState;
}

export type SPFxRuntimeListener = () => void;

export type SPFxRuntimeUpdater =
  | Partial<SPFxRuntimeState>
  | ((previous: SPFxRuntimeState) => SPFxRuntimeState);

export interface SPFxRuntimeStore {
  getState(): SPFxRuntimeState;
  setState(updater: SPFxRuntimeUpdater): void;
  subscribe(listener: SPFxRuntimeListener): () => void;
}

export function createDefaultSPFxRuntimeState(): SPFxRuntimeState {
  return {
    theme: undefined,
    displayMode: undefined,
    properties: undefined,
    containerEl: undefined,
    containerSize: undefined,
    teams: { supported: false, initialized: false },
  };
}

function mergeState(
  previous: SPFxRuntimeState,
  partial: Partial<SPFxRuntimeState>
): SPFxRuntimeState {
  let changed = false;
  const next: SPFxRuntimeState = {
    ...previous,
    ...partial,
  };

  for (const key of Object.keys(next) as Array<keyof SPFxRuntimeState>) {
    if (!Object.is(previous[key], next[key])) {
      changed = true;
      break;
    }
  }

  return changed ? next : previous;
}

export function createSPFxRuntimeStore(
  initialState?: Partial<SPFxRuntimeState>
): SPFxRuntimeStore {
  let state = mergeState(createDefaultSPFxRuntimeState(), initialState ?? {});
  const listeners = new Set<SPFxRuntimeListener>();

  const notify = (): void => {
    for (const listener of Array.from(listeners)) {
      listener();
    }
  };

  return {
    getState: () => state,

    setState: (updater: SPFxRuntimeUpdater): void => {
      const nextState = typeof updater === 'function'
        ? updater(state)
        : mergeState(state, updater);

      if (Object.is(nextState, state)) {
        return;
      }

      state = nextState;

      notify();
    },

    subscribe: (listener: SPFxRuntimeListener): (() => void) => {
      let subscribed = true;
      listeners.add(listener);

      return () => {
        if (!subscribed) {
          return;
        }

        subscribed = false;
        listeners.delete(listener);
      };
    },
  };
}
