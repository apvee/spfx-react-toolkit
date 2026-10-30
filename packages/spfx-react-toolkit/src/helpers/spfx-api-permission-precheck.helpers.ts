/**
 * Token acquisition policy used by the precheck service. `passive` disables provider interaction; `interactiveAllowed` permits it. Neither mode grants permissions or bypasses consent.
 *
 * @example
 * ```ts
 * const mode: SPFxApiPermissionPrecheckMode = 'passive';
 * ```
 */
export type SPFxApiPermissionPrecheckMode = 'passive' | 'interactiveAllowed';

/**
 * Aggregate UI state: `idle` has no results, `checking` has pending work, `ready` has every required scope, `actionRequired` has remediable required failures, and `cannotDetermine` has unresolved required results.
 *
 * @example
 * ```ts
 * const state = summarizeSPFxApiPermissionResults([]).configurationState;
 * ```
 */
export type SPFxApiPermissionConfigurationState =
  | 'idle'
  | 'checking'
  | 'ready'
  | 'actionRequired'
  | 'cannotDetermine';

/**
 * Outcome of one delegated-scope precheck. `available` means the decoded token contains an accepted scope; other values describe configuration, token or acquisition failures. This is diagnostic evidence, not proof that an API operation will succeed.
 *
 * @example
 * ```ts
 * const status: SPFxApiPermissionCheckStatus = 'missingScope';
 * ```
 */
export type SPFxApiPermissionCheckStatus =
  | 'available'
  | 'missingScope'
  | 'consentRequired'
  | 'interactionRequired'
  | 'claimsChallenge'
  | 'conditionalAccessBlocked'
  | 'accessRestricted'
  | 'resourceMisconfigured'
  | 'tokenProviderUnavailable'
  | 'audienceMismatch'
  | 'unsupportedToken'
  | 'unsupportedPermissionKind'
  | 'tokenExpired'
  | 'transientFailure'
  | 'timeout'
  | 'invalidRequirement'
  | 'unknown';

/**
 * Presentation severity: available is ok; unknown, transientFailure and timeout are unknown regardless of requiredness; all other unavailable statuses are error for required requirements and warning for optional ones. This differs from aggregate readiness buckets.
 *
 * @example
 * ```ts
 * const severity: SPFxApiPermissionSeverity = 'warning';
 * ```
 */
export type SPFxApiPermissionSeverity = 'ok' | 'warning' | 'error' | 'unknown';

/**
 * One delegated scope with optional alternatives and UI metadata. Scope names are case-sensitive; normalization trims whitespace and defaults `required` to true.
 *
 * @example
 * ```ts
 * const scope: SPFxPermissionScopeOptions = { scope: 'User.Read', required: false };
 * ```
 */
export interface SPFxPermissionScopeOptions {
  /**
   * Case-sensitive delegated scope name; normalization trims surrounding whitespace.
   */
  readonly scope: string;
  /**
   * Whether this scope blocks configuration readiness. Defaults to true when omitted.
   */
  readonly required?: boolean;
  /**
   * Explicit alternative scope names accepted as equivalent; no hierarchy is inferred. Omitted means no alternatives.
   */
  readonly satisfies?: readonly string[];
  /**
   * Optional display label preserved on the normalized requirement.
   */
  readonly label?: string;
}

/**
 * Convenient scope input: a string is a required scope; an options object specifies optionality, acceptable alternative scopes and a display label.
 *
 * @example
 * ```ts
 * const scopes: SPFxPermissionScopeInput[] = ['User.Read', { scope: 'Mail.Read', required: false }];
 * ```
 */
export type SPFxPermissionScopeInput = string | SPFxPermissionScopeOptions;

/**
 * Custom API resource and its delegated scopes. `resource` is the token request resource; `packageResource` is the SharePoint API-access resource name. Audience checks are skipped unless expected audiences are supplied.
 *
 * @example
 * ```ts
 * const api: SPFxCustomApiPermissionInput = { name: 'Contoso API', resource: 'api://contoso', scopes: ['access_as_user'] };
 * ```
 */
export interface SPFxCustomApiPermissionInput {
  /**
   * Optional identity prefix. Defaults to the resource endpoint when omitted or empty.
   */
  readonly id?: string;
  /**
   * Human-readable resource name; also the default package resource label.
   */
  readonly name: string;
  /**
   * Resource URI passed to the SPFx token provider.
   */
  readonly resource: string;
  /**
   * Resource label for SharePoint permission requests; defaults to name when omitted or empty.
   */
  readonly packageResource?: string;
  /**
   * Delegated scope inputs to normalize; each produces one requirement.
   */
  readonly scopes: readonly SPFxPermissionScopeInput[];
  /**
   * Exact accepted token audience strings. Omission disables audience comparison for this API.
   */
  readonly expectedAudiences?: readonly string[];
}

/**
 * Inputs normalized into delegated-scope requirements, in Graph, custom API, then explicit-requirement order. Duplicate endpoint/scope pairs keep the first entry. Omitted arrays contribute no requirements; normalization performs no token request.
 *
 * @example
 * ```ts
 * const config: SPFxApiPermissionPrecheckConfig = { graph: ['User.Read'] };
 * ```
 */
export interface SPFxApiPermissionPrecheckConfig {
  /**
   * Graph delegated scopes; omitted means no Graph requirements.
   */
  readonly graph?: readonly SPFxPermissionScopeInput[];
  /**
   * Custom resource configurations; omitted means no custom API requirements.
   */
  readonly customApis?: readonly SPFxCustomApiPermissionInput[];
  /**
   * Explicit normalized requirements appended after configured resources. Duplicate endpoint/scope pairs keep the earlier requirement.
   */
  readonly requirements?: readonly SPFxApiPermissionRequirement[];
}

/**
 * Normalized, required metadata for one delegated-scope check. Prefer `normalizeSPFxApiPermissionRequirements` to construct Graph/custom API entries consistently. Validation checks completeness, not endpoint reachability or consent.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * ```
 */
export interface SPFxApiPermissionRequirement {
  /**
   * Stable identity used to associate a result with its requirement; must be nonblank.
   */
  readonly id: string;
  /**
   * Resource display name; must be nonblank.
   */
  readonly resourceName: string;
  /**
   * Token request resource URI; must be nonblank.
   */
  readonly resourceEndpoint: string;
  /**
   * Resource label in package-solution API permission requests; must be nonblank.
   */
  readonly packageResource: string;
  /**
   * Case-sensitive delegated scope name; must be nonblank.
   */
  readonly scope: string;
  /**
   * Supported permission kind. Only delegated scopes are evaluated, not application roles.
   */
  readonly kind: 'delegatedScope';
  /**
   * Whether an unavailable result blocks aggregate readiness.
   */
  readonly required: boolean;
  /**
   * Explicit alternative accepted scope names; omitted means only scope is accepted.
   */
  readonly satisfies?: readonly string[];
  /**
   * Exact accepted audience strings; omitted or empty disables audience comparison.
   */
  readonly expectedAudiences?: readonly string[];
  /**
   * Administrator remediation text displayed on actionable failures.
   */
  readonly adminMessage: string;
  /**
   * Optional application-defined display label.
   */
  readonly label?: string;
}

/**
 * Resource/scope pair suitable for `webApiPermissionRequests` in package-solution.json. Listing a request does not approve it; an administrator must approve API access.
 *
 * @example
 * ```ts
 * const request: SPFxApiPermissionPackageSolutionEntry = { resource: 'Microsoft Graph', scope: 'User.Read' };
 * ```
 */
export interface SPFxApiPermissionPackageSolutionEntry {
  /**
   * Resource label for a package-solution.json API access request.
   */
  readonly resource: string;
  /**
   * Delegated permission scope requested from that resource.
   */
  readonly scope: string;
}

/**
 * Token-free display metadata for a checked requirement, including diagnostic status and administrator remediation. Consumers should branch on `status`, rather than infer success from a message or severity.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * const summary: SPFxApiPermissionSummary = evaluateSPFxApiPermissionRequirement(requirement, { scp: 'User.Read' });
 * ```
 */
export interface SPFxApiPermissionSummary {
  /**
   * Identity copied from the checked requirement.
   */
  readonly id: string;
  /**
   * Resource display name copied from the requirement.
   */
  readonly resourceName: string;
  /**
   * Token request resource URI copied from the requirement.
   */
  readonly resourceEndpoint: string;
  /**
   * Requested delegated scope, in its original case.
   */
  readonly scope: string;
  /**
   * Whether this requirement blocks readiness when unavailable.
   */
  readonly required: boolean;
  /**
   * Diagnostic outcome for this requirement; inspect available for scope success.
   */
  readonly status: SPFxApiPermissionCheckStatus;
  /**
   * UI severity derived from status and whether the requirement is required.
   */
  readonly severity: SPFxApiPermissionSeverity;
  /**
   * Human-readable diagnostic text, without the access token.
   */
  readonly message: string;
  /**
   * Administrator remediation instruction from the requirement.
   */
  readonly adminMessage: string;
  /**
   * Resource/scope pair to include in package-solution.json; approval remains separate.
   */
  readonly packageSolutionEntry: SPFxApiPermissionPackageSolutionEntry;
  /**
   * Requested or alternative scope found exactly in the token; undefined when none matches.
   */
  readonly matchedScope?: string;
  /**
   * Extracted acquisition error code when available; undefined otherwise.
   */
  readonly errorCode?: string;
}

/**
 * Detailed precheck outcome extending `SPFxApiPermissionSummary` with observed delegated scopes, audience and tenant metadata. Raw access tokens are never included; these claims are diagnostic and unverified.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * const result: SPFxApiPermissionCheckResult = evaluateSPFxApiPermissionRequirement(requirement, { scp: 'User.Read' });
 * ```
 */
export interface SPFxApiPermissionCheckResult extends SPFxApiPermissionSummary {
  /**
   * Delegated scp claim values observed in the token; empty when no scopes were decoded.
   */
  readonly detectedScopes: readonly string[];
  /**
   * First string audience observed when available; undefined for absent/unusable audiences.
   */
  readonly audience?: string;
  /**
   * Tenant ID claim when available; not independently verified.
   */
  readonly tenantId?: string;
  /**
   * Observed scope differing only by case from an acceptable scope; diagnostic only, not accepted.
   */
  readonly caseMismatchScope?: string;
}

/**
 * Decoded JWT claims used for diagnostics. All known claims are optional and extra claims are retained as unknown values. Decoding does not validate signatures or claim types; never use this object as authorization evidence.
 *
 * @example
 * ```ts
 * const payload: SPFxJwtPayload = { scp: 'User.Read Mail.Read', aud: SPFX_GRAPH_APP_ID };
 * ```
 */
export interface SPFxJwtPayload {
  /**
   * Optional resource audience string or list; decoded without signature verification.
   */
  readonly aud?: string | readonly string[];
  /**
   * Optional space-separated delegated scope claim.
   */
  readonly scp?: string;
  /**
   * Optional application role claims; delegated-scope prechecks do not accept these as scopes.
   */
  readonly roles?: readonly string[] | string;
  /**
   * Optional tenant identifier claim; unverified.
   */
  readonly tid?: string;
  /**
   * Optional expiration time in Unix seconds; numeric values are checked by evaluation.
   */
  readonly exp?: number;
  /**
   * Optional earliest valid time in Unix seconds; numeric values are checked by evaluation.
   */
  readonly nbf?: number;
  /** Additional decoded claims; callers must validate their types before use. */
  readonly [key: string]: unknown;
}

/**
 * Microsoft Graph resource label used in SharePoint package-solution permission requests.
 *
 * @example
 * ```ts
 * const request = { resource: SPFX_GRAPH_RESOURCE_NAME, scope: 'User.Read' };
 * ```
 */
export const SPFX_GRAPH_RESOURCE_NAME = 'Microsoft Graph';
/**
 * Microsoft Graph resource endpoint used to acquire tokens and construct normalized Graph requirement IDs.
 *
 * @example
 * ```ts
 * const id = createSPFxApiPermissionRequirementId(SPFX_GRAPH_RESOURCE_ENDPOINT, 'User.Read');
 * ```
 */
export const SPFX_GRAPH_RESOURCE_ENDPOINT = 'https://graph.microsoft.com';
/**
 * Microsoft Graph application audience accepted alongside its resource endpoint by normalized Graph requirements.
 *
 * @example
 * ```ts
 * const payload: SPFxJwtPayload = { aud: SPFX_GRAPH_APP_ID, scp: 'User.Read' };
 * ```
 */
export const SPFX_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

const graphExpectedAudiences: readonly string[] = [
  SPFX_GRAPH_RESOURCE_ENDPOINT,
  SPFX_GRAPH_APP_ID
];

const actionRequiredStatuses: readonly SPFxApiPermissionCheckStatus[] = [
  'missingScope',
  'consentRequired',
  'interactionRequired',
  'claimsChallenge',
  'conditionalAccessBlocked',
  'accessRestricted'
];

function normalizeScopeInput(scopeInput: SPFxPermissionScopeInput): {
  readonly scope: string;
  readonly required: boolean;
  readonly satisfies?: readonly string[];
  readonly label?: string;
} {
  if (typeof scopeInput === 'string') {
    return {
      scope: scopeInput.trim(),
      required: true
    };
  }

  const satisfies = scopeInput.satisfies
    ?.map(scope => scope.trim())
    .filter(scope => scope.length > 0);

  return {
    scope: scopeInput.scope.trim(),
    required: scopeInput.required !== false,
    satisfies: satisfies && satisfies.length > 0 ? satisfies : undefined,
    label: scopeInput.label
  };
}

function getDeduplicationKey(requirement: SPFxApiPermissionRequirement): string {
  return `${requirement.resourceEndpoint}\n${requirement.scope}`;
}

function addRequirement(
  requirements: SPFxApiPermissionRequirement[],
  seen: Set<string>,
  requirement: SPFxApiPermissionRequirement
): void {
  const key = getDeduplicationKey(requirement);

  if (seen.has(key)) {
    return;
  }

  seen.add(key);
  requirements.push(requirement);
}

/**
 * Normalizes configured Graph, custom API and explicit requirements without I/O. Trims configured scope names; Graph and custom scopes default to required. Deduplicates exact endpoint/scope pairs, keeping the first entry; explicit requirements are otherwise preserved.
 *
 * @param config - Configuration; omitted arrays contribute no requirements.
 * @returns Ordered, deduplicated requirements; an empty configuration returns an empty array.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * ```
 */
export function normalizeSPFxApiPermissionRequirements(
  config: SPFxApiPermissionPrecheckConfig
): readonly SPFxApiPermissionRequirement[] {
  const requirements: SPFxApiPermissionRequirement[] = [];
  const seen = new Set<string>();

  for (const scopeInput of config.graph || []) {
    const normalizedScope = normalizeScopeInput(scopeInput);

    addRequirement(requirements, seen, {
      id: createSPFxApiPermissionRequirementId(SPFX_GRAPH_RESOURCE_ENDPOINT, normalizedScope.scope, 'graph'),
      resourceName: SPFX_GRAPH_RESOURCE_NAME,
      resourceEndpoint: SPFX_GRAPH_RESOURCE_ENDPOINT,
      packageResource: SPFX_GRAPH_RESOURCE_NAME,
      scope: normalizedScope.scope,
      kind: 'delegatedScope',
      required: normalizedScope.required,
      satisfies: normalizedScope.satisfies,
      expectedAudiences: graphExpectedAudiences,
      adminMessage: buildSPFxApiPermissionAdminMessage(SPFX_GRAPH_RESOURCE_NAME, normalizedScope.scope),
      label: normalizedScope.label
    });
  }

  for (const customApi of config.customApis || []) {
    for (const scopeInput of customApi.scopes || []) {
      const normalizedScope = normalizeScopeInput(scopeInput);
      const packageResource = customApi.packageResource || customApi.name;

      addRequirement(requirements, seen, {
        id: createSPFxApiPermissionRequirementId(
          customApi.resource,
          normalizedScope.scope,
          customApi.id || customApi.resource
        ),
        resourceName: customApi.name,
        resourceEndpoint: customApi.resource,
        packageResource,
        scope: normalizedScope.scope,
        kind: 'delegatedScope',
        required: normalizedScope.required,
        satisfies: normalizedScope.satisfies,
        expectedAudiences: customApi.expectedAudiences,
        adminMessage: buildSPFxApiPermissionAdminMessage(packageResource, normalizedScope.scope),
        label: normalizedScope.label
      });
    }
  }

  for (const requirement of config.requirements || []) {
    addRequirement(requirements, seen, requirement);
  }

  return requirements;
}

/**
 * Builds a display-only administrator instruction for one resource/scope pair. It does not change package configuration or grant API permissions.
 *
 * @param packageResource - Resource label used in package-solution.json.
 * @param scope - Delegated scope to approve.
 * @returns Plain-text instruction for SharePoint Admin Center API access.
 *
 * @example
 * ```ts
 * const message = buildSPFxApiPermissionAdminMessage('Microsoft Graph', 'User.Read');
 * ```
 */
export function buildSPFxApiPermissionAdminMessage(
  packageResource: string,
  scope: string
): string {
  return `Approve ${packageResource} / ${scope} in SharePoint Admin Center API access.`;
}

/**
 * Builds a deterministic display identity. Graph uses `graph:<trimmed scope>`; other resources use `api:<trimmed prefix or endpoint>:<trimmed scope>`. Values are not URL-encoded.
 *
 * @param resourceEndpoint - Token resource endpoint.
 * @param scope - Delegated scope; surrounding whitespace is removed.
 * @param prefix - Optional custom namespace; `graph` selects the Graph identity format.
 * @returns Requirement identity string.
 *
 * @example
 * ```ts
 * const id = createSPFxApiPermissionRequirementId(SPFX_GRAPH_RESOURCE_ENDPOINT, 'User.Read');
 * ```
 */
export function createSPFxApiPermissionRequirementId(
  resourceEndpoint: string,
  scope: string,
  prefix?: string
): string {
  const trimmedScope = scope.trim();

  if (prefix === 'graph' || (!prefix && resourceEndpoint === SPFX_GRAPH_RESOURCE_ENDPOINT)) {
    return `graph:${trimmedScope}`;
  }

  return `api:${(prefix || resourceEndpoint).trim()}:${trimmedScope}`;
}

/**
 * Checks required metadata for a delegated-scope requirement without acquiring a token. Unsupported kinds return `unsupportedPermissionKind`; missing or blank required fields return `invalidRequirement`.
 *
 * @param requirement - Requirement to inspect, including metadata and permission kind.
 * @returns A diagnostic failure result, or undefined when metadata is complete.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * const issue = validateSPFxApiPermissionRequirement(requirement);
 * ```
 */
export function validateSPFxApiPermissionRequirement(
  requirement: SPFxApiPermissionRequirement
): SPFxApiPermissionCheckResult | undefined {
  const kind = (requirement as { readonly kind?: string }).kind;

  if (kind && kind !== 'delegatedScope') {
    return createSPFxApiPermissionResult(
      requirement,
      'unsupportedPermissionKind',
      'This precheck only supports delegated permission scopes.'
    );
  }

  if (
    !requirement.id?.trim() ||
    !requirement.resourceName?.trim() ||
    !requirement.resourceEndpoint?.trim() ||
    !requirement.packageResource?.trim() ||
    !requirement.scope?.trim() ||
    !kind
  ) {
    return createSPFxApiPermissionResult(
      requirement,
      'invalidRequirement',
      'The API permission requirement is incomplete.'
    );
  }

  return undefined;
}

/**
 * Decodes only the JSON payload of a JWT-shaped string. Malformed/nonobject payloads return undefined; signature, issuer and claim types are not validated. Do not log raw tokens or treat decoded claims as verified identity.
 *
 * @param token - Token text to decode locally; no network request is made.
 * @returns Decoded object claims, or undefined when decoding fails.
 *
 * @example
 * ```ts
 * const payload = decodeSPFxJwtPayload('header.eyJzY3AiOiJVc2VyLlJlYWQifQ.signature');
 * ```
 */
export function decodeSPFxJwtPayload(token: string): SPFxJwtPayload | undefined {
  try {
    const parts = token.split('.');

    if (parts.length < 3 || !parts[1]) {
      return undefined;
    }

    const payloadJson = decodeBase64Url(parts[1]);
    const payload = JSON.parse(payloadJson);

    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      return undefined;
    }

    return payload as SPFxJwtPayload;
  } catch {
    return undefined;
  }
}

/**
 * Extracts whitespace-trimmed scopes from a string `scp` claim. Ignores app roles and returns an empty array for absent/nonstring claims; duplicate scopes are retained.
 *
 * @param payload - Decoded, unverified JWT claims.
 * @returns Delegated scope names in their original case and order.
 *
 * @example
 * ```ts
 * const scopes = extractSPFxDelegatedScopes({ scp: 'User.Read Mail.Read' });
 * ```
 */
export function extractSPFxDelegatedScopes(payload: SPFxJwtPayload): readonly string[] {
  if (typeof payload.scp !== 'string') {
    return [];
  }

  return payload.scp
    .split(' ')
    .map(scope => scope.trim())
    .filter(scope => scope.length > 0);
}

/**
 * Evaluates a requirement against decoded claims without I/O or signature verification. Checks metadata, token dates, optional audience policy and exact case-sensitive delegated scopes (including explicitly accepted alternatives). App-only roles are unsupported. An available result does not prove an API request will succeed.
 *
 * @param requirement - Normalized delegated-scope requirement.
 * @param payload - Decoded claims; undefined produces `unsupportedToken`.
 * @param options - Audience validation is enabled unless `validateAudience` is false; absent expected audiences skip the comparison.
 * @returns Detailed diagnostic result without a raw token.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * const result = evaluateSPFxApiPermissionRequirement(requirement, { aud: SPFX_GRAPH_APP_ID, scp: 'User.Read' });
 * ```
 */
export function evaluateSPFxApiPermissionRequirement(
  requirement: SPFxApiPermissionRequirement,
  payload: SPFxJwtPayload | undefined,
  options?: { /**
   * Checks configured expected audiences unless false. Omitted defaults to enabled.
   */
  readonly validateAudience?: boolean }
): SPFxApiPermissionCheckResult {
  const invalidResult = validateSPFxApiPermissionRequirement(requirement);

  if (invalidResult) {
    return invalidResult;
  }

  if (!payload) {
    return createSPFxApiPermissionResult(
      requirement,
      'unsupportedToken',
      'The access token could not be decoded as a JWT.'
    );
  }

  const audience = getAudience(payload);
  const detectedScopes = extractSPFxDelegatedScopes(payload);
  const baseDetails = {
    audience,
    detectedScopes,
    tenantId: payload.tid
  };

  if (isTokenExpiredOrNotYetValid(payload)) {
    return createSPFxApiPermissionResult(
      requirement,
      'tokenExpired',
      'The access token is expired or is not valid yet.',
      baseDetails
    );
  }

  if (options?.validateAudience !== false && hasAudienceMismatch(requirement, payload)) {
    return createSPFxApiPermissionResult(
      requirement,
      'audienceMismatch',
      `The access token audience does not match ${requirement.resourceName}.`,
      baseDetails
    );
  }

  if (detectedScopes.length === 0 && hasRoles(payload)) {
    return createSPFxApiPermissionResult(
      requirement,
      'unsupportedPermissionKind',
      'The token contains app roles, but this precheck validates delegated scp scopes only.',
      baseDetails
    );
  }

  const acceptableScopes = [
    requirement.scope,
    ...(requirement.satisfies || [])
  ];
  const matchedScope = acceptableScopes.find(scope => detectedScopes.indexOf(scope) >= 0);

  if (matchedScope) {
    return createSPFxApiPermissionResult(
      requirement,
      'available',
      `${requirement.resourceName} / ${requirement.scope} is available to the current SPFx runtime.`,
      {
        ...baseDetails,
        matchedScope
      }
    );
  }

  const acceptableLowerScopes = acceptableScopes.map(scope => scope.toLowerCase());
  const caseMismatchScope = detectedScopes.find(
    scope => acceptableLowerScopes.indexOf(scope.toLowerCase()) >= 0
  );

  return createSPFxApiPermissionResult(
    requirement,
    'missingScope',
    `${requirement.resourceName} / ${requirement.scope} is not present in the delegated token scopes.`,
    {
      ...baseDetails,
      caseMismatchScope
    }
  );
}

/**
 * Classifies a token-acquisition error using message/code heuristics into consent, claims, Conditional Access, interaction, assignment, resource, timeout, transient or unknown outcomes. Classification is diagnostic and does not retry or grant access.
 *
 * @param requirement - Requirement whose display/remediation metadata accompanies the failure.
 * @param error - Arbitrary thrown value from token acquisition.
 * @returns Diagnostic result with an extracted error code when available.
 *
 * @example
 * ```ts
 * const [requirement] = normalizeSPFxApiPermissionRequirements({ graph: ['User.Read'] });
 * const result = classifySPFxApiPermissionError(requirement, new Error('AADSTS65001: consent required'));
 * ```
 */
export function classifySPFxApiPermissionError(
  requirement: SPFxApiPermissionRequirement,
  error: unknown
): SPFxApiPermissionCheckResult {
  const errorText = getErrorText(error);
  const normalizedErrorText = errorText.toLowerCase();
  const errorCode = getErrorCode(error, errorText);

  if (normalizedErrorText.indexOf('aadsts65001') >= 0 || normalizedErrorText.indexOf('consent') >= 0) {
    return createSPFxApiPermissionResult(
      requirement,
      'consentRequired',
      `${requirement.adminMessage}`,
      { errorCode }
    );
  }

  if (normalizedErrorText.indexOf('claims') >= 0) {
    return createSPFxApiPermissionResult(
      requirement,
      'claimsChallenge',
      'A claims challenge is required before this token can be acquired.',
      { errorCode }
    );
  }

  if (normalizedErrorText.indexOf('conditional access') >= 0) {
    return createSPFxApiPermissionResult(
      requirement,
      'conditionalAccessBlocked',
      'Conditional Access blocked token acquisition for this API permission.',
      { errorCode }
    );
  }

  if (
    normalizedErrorText.indexOf('login') >= 0 ||
    normalizedErrorText.indexOf('mfa') >= 0 ||
    normalizedErrorText.indexOf('interaction') >= 0 ||
    normalizedErrorText.indexOf('interactive') >= 0 ||
    normalizedErrorText.indexOf('authentication popup') >= 0 ||
    (normalizedErrorText.indexOf('passive') >= 0 && normalizedErrorText.indexOf('popup') >= 0)
  ) {
    return createSPFxApiPermissionResult(
      requirement,
      'interactionRequired',
      'User interaction is required before this token can be acquired.',
      { errorCode }
    );
  }

  if (normalizedErrorText.indexOf('user assignment') >= 0 || normalizedErrorText.indexOf('access assignment') >= 0) {
    return createSPFxApiPermissionResult(
      requirement,
      'accessRestricted',
      'Access to this API is restricted by assignment.',
      { errorCode }
    );
  }

  if (normalizedErrorText.indexOf('invalid resource') >= 0 || normalizedErrorText.indexOf('not found') >= 0) {
    return createSPFxApiPermissionResult(
      requirement,
      'resourceMisconfigured',
      'The API resource could not be resolved by the token provider.',
      { errorCode }
    );
  }

  if (normalizedErrorText.indexOf('timeout') >= 0 || normalizedErrorText.indexOf('timed out') >= 0) {
    return createSPFxApiPermissionResult(
      requirement,
      'timeout',
      'Token acquisition timed out.',
      { errorCode }
    );
  }

  if (
    normalizedErrorText.indexOf('network') >= 0 ||
    normalizedErrorText.indexOf('throttle') >= 0 ||
    normalizedErrorText.indexOf('temporarily unavailable') >= 0 ||
    normalizedErrorText.indexOf('temporary') >= 0
  ) {
    return createSPFxApiPermissionResult(
      requirement,
      'transientFailure',
      'A transient token acquisition failure occurred.',
      { errorCode }
    );
  }

  return createSPFxApiPermissionResult(
    requirement,
    'unknown',
    'The API permission state could not be determined.',
    { errorCode }
  );
}

/**
 * Groups available results, required actionable failures, optional failures and unresolved required outcomes. `isConfigured` checks only required results, so an empty set returns true even though its configuration state is `idle`. A checking flag overrides the aggregate state.
 *
 * @param results - Completed requirement results.
 * @param isChecking - Whether work is still pending; omitted means false.
 * @returns Aggregate state and token-free UI buckets.
 *
 * @example
 * ```ts
 * const summary = summarizeSPFxApiPermissionResults([]);
 * // summary.configurationState === 'idle'; summary.isConfigured === true
 * ```
 */
export function summarizeSPFxApiPermissionResults(
  results: readonly SPFxApiPermissionCheckResult[],
  isChecking?: boolean
): {
  /**
   * Aggregate readiness state; checking takes priority, while no results produces idle.
   */
  readonly configurationState: SPFxApiPermissionConfigurationState;
  /**
   * True when every required result is available. Empty/no-required results also yield true.
   */
  readonly isConfigured: boolean;
  /**
   * All results with available status, whether required or optional.
   */
  readonly available: readonly SPFxApiPermissionSummary[];
  /**
   * Required failures whose status indicates an actionable permission/access issue.
   */
  readonly missing: readonly SPFxApiPermissionSummary[];
  /**
   * All unavailable optional requirements; they do not block isConfigured.
   */
  readonly warnings: readonly SPFxApiPermissionSummary[];
  /**
   * Required failures whose status does not determine an actionable missing permission.
   */
  readonly unknown: readonly SPFxApiPermissionSummary[];
} {
  const available: SPFxApiPermissionSummary[] = [];
  const missing: SPFxApiPermissionSummary[] = [];
  const warnings: SPFxApiPermissionSummary[] = [];
  const unknown: SPFxApiPermissionSummary[] = [];

  for (const result of results) {
    if (result.status === 'available') {
      available.push(result);
    } else if (!result.required) {
      warnings.push(result);
    } else if (isActionRequiredStatus(result.status)) {
      missing.push(result);
    } else {
      unknown.push(result);
    }
  }

  const requiredResults = results.filter(result => result.required);
  const isConfigured = requiredResults.every(result => result.status === 'available');
  let configurationState: SPFxApiPermissionConfigurationState;

  if (isChecking) {
    configurationState = 'checking';
  } else if (results.length === 0) {
    configurationState = 'idle';
  } else if (missing.length > 0) {
    configurationState = 'actionRequired';
  } else if (isConfigured) {
    configurationState = 'ready';
  } else {
    configurationState = 'cannotDetermine';
  }

  return {
    configurationState,
    isConfigured,
    available,
    missing,
    warnings,
    unknown
  };
}

function createSPFxApiPermissionResult(
  requirement: SPFxApiPermissionRequirement,
  status: SPFxApiPermissionCheckStatus,
  message: string,
  details?: {
    readonly detectedScopes?: readonly string[];
    readonly audience?: string;
    readonly tenantId?: string;
    readonly matchedScope?: string;
    readonly caseMismatchScope?: string;
    readonly errorCode?: string;
  }
): SPFxApiPermissionCheckResult {
  return {
    id: requirement.id || '',
    resourceName: requirement.resourceName || '',
    resourceEndpoint: requirement.resourceEndpoint || '',
    scope: requirement.scope || '',
    required: requirement.required !== false,
    status,
    severity: getSeverity(requirement, status),
    message,
    adminMessage: requirement.adminMessage || buildSPFxApiPermissionAdminMessage(
      requirement.packageResource || requirement.resourceName || requirement.resourceEndpoint || 'API',
      requirement.scope || ''
    ),
    packageSolutionEntry: {
      resource: requirement.packageResource || requirement.resourceName || requirement.resourceEndpoint || '',
      scope: requirement.scope || ''
    },
    detectedScopes: details?.detectedScopes || [],
    audience: details?.audience,
    tenantId: details?.tenantId,
    matchedScope: details?.matchedScope,
    caseMismatchScope: details?.caseMismatchScope,
    errorCode: details?.errorCode
  };
}

function getSeverity(
  requirement: SPFxApiPermissionRequirement,
  status: SPFxApiPermissionCheckStatus
): SPFxApiPermissionSeverity {
  if (status === 'available') {
    return 'ok';
  }

  if (status === 'unknown' || status === 'transientFailure' || status === 'timeout') {
    return 'unknown';
  }

  return requirement.required === false ? 'warning' : 'error';
}

function decodeBase64Url(value: string): string {
  const base64 = value
    .replace(/-/g, '+')
    .replace(/_/g, '/')
    .padEnd(Math.ceil(value.length / 4) * 4, '=');
  const binary = globalThis.atob(base64);
  const encoded = Array.prototype.map.call(
    binary,
    (character: string) => `%${character.charCodeAt(0).toString(16).padStart(2, '0')}`
  ).join('');

  return decodeURIComponent(encoded);
}

function getAudience(payload: SPFxJwtPayload): string | undefined {
  if (typeof payload.aud === 'string') {
    return payload.aud;
  }

  if (Array.isArray(payload.aud) && typeof payload.aud[0] === 'string') {
    return payload.aud[0];
  }

  return undefined;
}

function getAudiences(payload: SPFxJwtPayload): readonly string[] {
  if (typeof payload.aud === 'string') {
    return [payload.aud];
  }

  if (Array.isArray(payload.aud)) {
    return payload.aud.filter((audience): audience is string => typeof audience === 'string');
  }

  return [];
}

function hasAudienceMismatch(
  requirement: SPFxApiPermissionRequirement,
  payload: SPFxJwtPayload
): boolean {
  if (!requirement.expectedAudiences || requirement.expectedAudiences.length === 0) {
    return false;
  }

  const tokenAudiences = getAudiences(payload);

  if (tokenAudiences.length === 0) {
    return true;
  }

  return !tokenAudiences.some(audience => requirement.expectedAudiences?.indexOf(audience) !== -1);
}

function isTokenExpiredOrNotYetValid(payload: SPFxJwtPayload): boolean {
  const now = Math.floor(Date.now() / 1000);

  return typeof payload.exp === 'number' && payload.exp <= now ||
    typeof payload.nbf === 'number' && payload.nbf > now;
}

function hasRoles(payload: SPFxJwtPayload): boolean {
  if (Array.isArray(payload.roles)) {
    return payload.roles.length > 0;
  }

  return typeof payload.roles === 'string' && payload.roles.length > 0;
}

function getErrorText(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === 'string') {
    return error;
  }

  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}

function getErrorCode(error: unknown, errorText: string): string | undefined {
  if (error && typeof error === 'object') {
    const code = (error as { readonly code?: unknown; readonly errorCode?: unknown }).code ||
      (error as { readonly code?: unknown; readonly errorCode?: unknown }).errorCode;

    if (typeof code === 'string' && code.length > 0) {
      return code;
    }
  }

  const aadMatch = /AADSTS\d+/i.exec(errorText);

  return aadMatch ? aadMatch[0].toUpperCase() : undefined;
}

function isActionRequiredStatus(status: SPFxApiPermissionCheckStatus): boolean {
  return actionRequiredStatuses.indexOf(status) >= 0;
}
