import type { IItems } from '@pnp/sp/items';

interface ProxyConstructor {
  new<T extends object>(
    target: T,
    handler: {
      get?(target: T, prop: string | symbol, receiver: unknown): unknown;
    }
  ): T;
}

declare const Proxy: ProxyConstructor;

export interface QueryTopTracker {
  top?: number;
}

export interface EffectiveListQuery {
  readonly query: IItems;
  readonly pageSize: number | undefined;
}

export function createMonitoredListQuery(
  target: IItems,
  tracker: QueryTopTracker
): IItems {
  return new Proxy(target, {
    get: function(t: IItems, prop: string | symbol): unknown {
      if (prop === 'top') {
        return function(n: number): IItems {
          tracker.top = n;
          const result = t.top.call(t, n);
          return createMonitoredListQuery(result as IItems, tracker);
        };
      }

      const value = (t as unknown as Record<string, unknown>)[prop as string];
      if (typeof value === 'function') {
        return function(...args: unknown[]): unknown {
          const result = value.apply(t, args);
          if (
            result &&
            typeof result === 'object' &&
            typeof (result as { select?: unknown }).select === 'function'
          ) {
            return createMonitoredListQuery(result as IItems, tracker);
          }
          return result;
        };
      }

      return value;
    }
  }) as IItems;
}

export function resolveEffectiveListQuery(
  userQuery: IItems,
  tracker: QueryTopTracker,
  pageSize: number | undefined
): EffectiveListQuery {
  if (tracker.top !== undefined) {
    return {
      query: userQuery,
      pageSize: tracker.top,
    };
  }

  if (pageSize !== undefined) {
    return {
      query: userQuery.top(pageSize),
      pageSize,
    };
  }

  return {
    query: userQuery,
    pageSize: undefined,
  };
}

export function hasMoreListItems(resultLength: number, pageSize: number | undefined): boolean {
  return pageSize !== undefined && resultLength === pageSize;
}
