import type { SPHttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import type { PickerValue } from './picker-values.types';
import type { PickerOption, PrincipalPickerInitialPrincipal, PickerSelectionEvent } from './picker.types';
import type { PickerRequestCache } from './picker-request-cache.internal';

export interface PickerSourceContext {
  readonly owner: object;
  readonly tenantId: string;
  readonly userId: string;
  readonly webUrl: string;
  readonly siteCollectionUrl: string;
  readonly spClient?: SPHttpClient;
  readonly graphClient?: MSGraphClientV3;
  /** Principal-only provider/factory lifetime, stable while its native Graph client initializes. */
  readonly graphClientIdentity?: object;
  /** Principal-only pending dependency; direct source contexts may omit it. */
  readonly getGraphClient?: () => Promise<MSGraphClientV3>;
}
export interface PickerSourceRequest {
  readonly query: string;
  readonly pageSize: number;
  readonly continuation?: string;
  readonly generation: number;
  readonly context: PickerSourceContext;
  readonly cache: PickerRequestCache;
  readonly partition?: string;
  readonly isCurrent?: () => boolean;
}
export interface PickerSourcePage<T> {
  readonly options: readonly T[];
  readonly continuation?: string;
  readonly completeness: 'complete' | 'partial';
  readonly errors?: readonly { readonly source: 'sharePoint' | 'graph'; readonly error: Error; /** False for bounded-result warnings that require narrowing criteria rather than retry. Omission means a recoverable source failure. */ readonly retryable?: boolean }[];
}
export interface PickerResolutionResult<V, T> {
  readonly resolved: readonly { readonly value: V; readonly option: T }[];
  readonly unresolved: readonly { readonly value: V; readonly error?: Error }[];
}
export type PickerHydrationInput<V> =
  | { readonly kind: 'value'; readonly key: string; readonly value: V }
  | { readonly kind: 'key'; readonly key: string }
  | { readonly kind: 'principalReference'; readonly key: string; readonly reference: PrincipalPickerInitialPrincipal };
export interface PickerHydrationResult<V, T> {
  readonly resolved: readonly { readonly key: string; readonly value: V; readonly option: T }[];
  readonly unresolved: readonly { readonly key: string; readonly label?: string; readonly error?: Error }[];
}
/** Shared lazy pre-search source data, including contextual Site and personal Principal suggestions. */
export interface PickerPersonalSuggestions<T> {
  readonly currentSite?: readonly T[];
  readonly hubSite?: readonly T[];
  readonly currentUser?: readonly T[];
  readonly people?: readonly T[];
  readonly groups?: readonly T[];
  readonly errors?: readonly { readonly source: 'people' | 'memberOf' | 'currentUser' | 'currentSite' | 'hubSite'; readonly error: Error }[];
  readonly completeness: 'complete' | 'partial';
}
export interface PickerSource<V extends PickerValue, T extends PickerOption<V>> {
  /** Changes only when transport identity or effective source parameters change. */
  readonly identity: string;
  /** Stable domain family for explicit invalidation of shared transport dependencies across configuration revisions. */
  readonly family?: string;
  /** Typed text is matched case-insensitively; query cache and continuations use deterministic lowercase. */
  readonly caseInsensitive?: boolean;
  /** Discovery returns the complete cached catalog, with typed matching performed locally. */
  readonly completeCatalog?: boolean;
  readonly context: PickerSourceContext;
  readonly ready: boolean;
  readonly transport: 'sharePoint' | 'graph';
  /** Consumer menu candidates require read-only server eligibility before display. */
  readonly preflight?: boolean;
  accepts(option: T): boolean;
  discover(request: PickerSourceRequest): Promise<PickerSourcePage<T>>;
  resolve(values: readonly V[], request: PickerSourceRequest): Promise<PickerResolutionResult<V, T>>;
  hydrate(inputs: readonly PickerHydrationInput<V>[], request: PickerSourceRequest): Promise<PickerHydrationResult<V, T>>;
  verifyCandidate(option: T, request: PickerSourceRequest): Promise<T | undefined>;
  /** Read-only check for IDs missing from the configured destination; never provisions. */
  needsIdentityResolution?(option: T, request: PickerSourceRequest): Promise<boolean>;
  /** Contextual candidates remain available during typed search, independently of the initial suggestion cap. */
  readonly persistentSuggestions?: boolean;
  personalSuggestions?(request: PickerSourceRequest): Promise<PickerPersonalSuggestions<T>>;
  confirm?(options: readonly T[], request: PickerSourceRequest, reason: 'selection' | 'explicitResolution'): Promise<readonly T[]>;
}
export interface PickerRuntimeAction {
  readonly event: PickerSelectionEvent;
  readonly key?: string;
}
