export type SPFxApiPermissionPrecheckMode = 'passive' | 'interactiveAllowed';

export type SPFxApiPermissionConfigurationState =
  | 'idle'
  | 'checking'
  | 'ready'
  | 'actionRequired'
  | 'cannotDetermine';

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

export type SPFxApiPermissionSeverity = 'ok' | 'warning' | 'error' | 'unknown';

export interface SPFxPermissionScopeOptions {
  readonly scope: string;
  readonly required?: boolean;
  readonly satisfies?: readonly string[];
  readonly label?: string;
}

export type SPFxPermissionScopeInput = string | SPFxPermissionScopeOptions;

export interface SPFxCustomApiPermissionInput {
  readonly id?: string;
  readonly name: string;
  readonly resource: string;
  readonly packageResource?: string;
  readonly scopes: readonly SPFxPermissionScopeInput[];
  readonly expectedAudiences?: readonly string[];
}

export interface SPFxApiPermissionPrecheckConfig {
  readonly graph?: readonly SPFxPermissionScopeInput[];
  readonly customApis?: readonly SPFxCustomApiPermissionInput[];
  readonly requirements?: readonly SPFxApiPermissionRequirement[];
}

export interface SPFxApiPermissionRequirement {
  readonly id: string;
  readonly resourceName: string;
  readonly resourceEndpoint: string;
  readonly packageResource: string;
  readonly scope: string;
  readonly kind: 'delegatedScope';
  readonly required: boolean;
  readonly satisfies?: readonly string[];
  readonly expectedAudiences?: readonly string[];
  readonly adminMessage: string;
  readonly label?: string;
}

export interface SPFxApiPermissionPackageSolutionEntry {
  readonly resource: string;
  readonly scope: string;
}

export interface SPFxApiPermissionSummary {
  readonly id: string;
  readonly resourceName: string;
  readonly resourceEndpoint: string;
  readonly scope: string;
  readonly required: boolean;
  readonly status: SPFxApiPermissionCheckStatus;
  readonly severity: SPFxApiPermissionSeverity;
  readonly message: string;
  readonly adminMessage: string;
  readonly packageSolutionEntry: SPFxApiPermissionPackageSolutionEntry;
  readonly matchedScope?: string;
  readonly errorCode?: string;
}

export interface SPFxApiPermissionCheckResult extends SPFxApiPermissionSummary {
  readonly detectedScopes: readonly string[];
  readonly audience?: string;
  readonly tenantId?: string;
  readonly caseMismatchScope?: string;
}

export interface SPFxJwtPayload {
  readonly aud?: string | readonly string[];
  readonly scp?: string;
  readonly roles?: readonly string[] | string;
  readonly tid?: string;
  readonly exp?: number;
  readonly nbf?: number;
  readonly [key: string]: unknown;
}

export const SPFX_GRAPH_RESOURCE_NAME = 'Microsoft Graph';
export const SPFX_GRAPH_RESOURCE_ENDPOINT = 'https://graph.microsoft.com';
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

export function buildSPFxApiPermissionAdminMessage(
  packageResource: string,
  scope: string
): string {
  return `Approve ${packageResource} / ${scope} in SharePoint Admin Center API access.`;
}

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
    !requirement.scope?.trim()
  ) {
    return createSPFxApiPermissionResult(
      requirement,
      'invalidRequirement',
      'The API permission requirement is incomplete.'
    );
  }

  return undefined;
}

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
  } catch (_error) {
    return undefined;
  }
}

export function extractSPFxDelegatedScopes(payload: SPFxJwtPayload): readonly string[] {
  if (typeof payload.scp !== 'string') {
    return [];
  }

  return payload.scp
    .split(' ')
    .map(scope => scope.trim())
    .filter(scope => scope.length > 0);
}

export function evaluateSPFxApiPermissionRequirement(
  requirement: SPFxApiPermissionRequirement,
  payload: SPFxJwtPayload | undefined,
  options?: { readonly validateAudience?: boolean }
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

  const acceptableLowerScopes = acceptableScopes.map(scope => scope.toLocaleLowerCase());
  const caseMismatchScope = detectedScopes.find(
    scope => acceptableLowerScopes.indexOf(scope.toLocaleLowerCase()) >= 0
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

export function classifySPFxApiPermissionError(
  requirement: SPFxApiPermissionRequirement,
  error: unknown
): SPFxApiPermissionCheckResult {
  const errorText = getErrorText(error);
  const normalizedErrorText = errorText.toLocaleLowerCase();
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
    normalizedErrorText.indexOf('interactive') >= 0
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

export function summarizeSPFxApiPermissionResults(
  results: readonly SPFxApiPermissionCheckResult[],
  isChecking?: boolean
): {
  readonly configurationState: SPFxApiPermissionConfigurationState;
  readonly isConfigured: boolean;
  readonly available: readonly SPFxApiPermissionSummary[];
  readonly missing: readonly SPFxApiPermissionSummary[];
  readonly warnings: readonly SPFxApiPermissionSummary[];
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
  } catch (_jsonError) {
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
