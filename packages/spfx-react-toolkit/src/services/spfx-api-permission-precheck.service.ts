import {
  classifySPFxApiPermissionError,
  decodeSPFxJwtPayload,
  evaluateSPFxApiPermissionRequirement,
  normalizeSPFxApiPermissionRequirements,
  validateSPFxApiPermissionRequirement,
  type SPFxApiPermissionCheckResult,
  type SPFxApiPermissionPrecheckConfig,
  type SPFxApiPermissionRequirement
} from '../helpers/spfx-api-permission-precheck.helpers';

/**
 * Minimal token acquisition contract accepted by permission diagnostics. Tokens are used locally and are not included in diagnostic results.
 *
 * @example
 * ```ts
 * const provider: SPFxAadTokenProviderLike = spfxTokenProvider;
 * ```
 */
export interface SPFxAadTokenProviderLike {
  /**
   * Acquires an access token; acquisition failures reject.
   * @param resourceEndpoint - Resource URL or application identifier.
   * @param options - Cache preference and optional claims passed to the provider.
   * @returns The access token.
   */
  getToken(
    resourceEndpoint: string,
    options?: {
      /** Whether the provider may reuse a cached token; omitted follows its native default. */
      readonly useCachedToken?: boolean;
      /** Optional claims challenge passed to the native provider. */
      readonly claims?: string
    }
  ): Promise<string>;
}

/**
 * Controls token acquisition and delegated-scope evaluation.
 *
 * @example
 * ```ts
 * const options: SPFxApiPermissionPrecheckServiceOptions = { useCachedToken: false, timeoutMs: 10000 };
 * ```
 */
export interface SPFxApiPermissionPrecheckServiceOptions {
  /**
   * Allows cached tokens unless false; defaults to true.
   */
  readonly useCachedToken?: boolean;
  /**
   * Checks expected audiences unless false; defaults to true.
   */
  readonly validateAudience?: boolean;
  /**
   * Acquisition timeout in milliseconds; defaults to 15000. Timing out does not cancel the provider request.
   */
  readonly timeoutMs?: number;
  /**
   * Acquires each resource group in sequence when true; otherwise acquires groups concurrently.
   */
  readonly sequentialResourceAcquisition?: boolean;
}

/**
 * Non-React delegated-scope diagnostics using one token acquisition per resource endpoint. Invalid requirements and acquisition failures become diagnostic results; this does not grant permissions.
 *
 * @example
 * ```ts
 * const service: SPFxApiPermissionPrecheckService = createSPFxApiPermissionPrecheckService(provider);
 * const results = await service.check({ graph: ['User.Read'] });
 * ```
 */
export interface SPFxApiPermissionPrecheckService {
  /**
   * Checks configured scopes without returning raw tokens.
   * @param config - Graph and custom API requirements.
   * @param options - Optional acquisition and audience checks.
   * @returns Per-scope diagnostics, including invalid requirements and token failures.
   */
  check(
    config: SPFxApiPermissionPrecheckConfig,
    options?: SPFxApiPermissionPrecheckServiceOptions
  ): Promise<readonly SPFxApiPermissionCheckResult[]>;
}

/**
 * Creates permission diagnostics without React state or provider context.
 * @param tokenProvider - Initialized token provider owned by the caller.
 * @returns A service that evaluates delegated scopes from token payloads.
 *
 * @example
 * ```ts
 * const service = createSPFxApiPermissionPrecheckService(provider);
 * const results = await service.check({ graph: ['User.Read'] });
 * ```
 */
export function createSPFxApiPermissionPrecheckService(
  tokenProvider: SPFxAadTokenProviderLike
): SPFxApiPermissionPrecheckService {
  return {
    check: (config, options) => checkSPFxApiPermissions(tokenProvider, config, options)
  };
}

async function checkSPFxApiPermissions(
  tokenProvider: SPFxAadTokenProviderLike,
  config: SPFxApiPermissionPrecheckConfig,
  options?: SPFxApiPermissionPrecheckServiceOptions
): Promise<readonly SPFxApiPermissionCheckResult[]> {
  const normalizedRequirements = normalizeSPFxApiPermissionRequirements(config);
  const validRequirements: SPFxApiPermissionRequirement[] = [];
  const results: SPFxApiPermissionCheckResult[] = [];

  for (const requirement of normalizedRequirements) {
    const invalidResult = validateSPFxApiPermissionRequirement(requirement);

    if (invalidResult) {
      results.push(invalidResult);
    } else {
      validRequirements.push(requirement);
    }
  }

  const requirementGroups = groupRequirementsByResourceEndpoint(validRequirements);

  if (options?.sequentialResourceAcquisition) {
    for (const requirements of requirementGroups.values()) {
      results.push(...await checkSPFxApiPermissionResourceGroup(tokenProvider, requirements, options));
    }

    return results;
  }

  const groupResults = await Promise.all(
    Array.from(requirementGroups.values()).map(requirements =>
      checkSPFxApiPermissionResourceGroup(tokenProvider, requirements, options)
    )
  );

  for (const resourceResults of groupResults) {
    results.push(...resourceResults);
  }

  return results;
}

async function checkSPFxApiPermissionResourceGroup(
  tokenProvider: SPFxAadTokenProviderLike,
  requirements: readonly SPFxApiPermissionRequirement[],
  options?: SPFxApiPermissionPrecheckServiceOptions
): Promise<readonly SPFxApiPermissionCheckResult[]> {
  const resourceEndpoint = requirements[0]?.resourceEndpoint;

  if (!resourceEndpoint) {
    return [];
  }

  try {
    const token = await getTokenWithTimeout(
      tokenProvider,
      resourceEndpoint,
      options?.useCachedToken !== false,
      options?.timeoutMs ?? 15000
    );
    const payload = decodeSPFxJwtPayload(token);

    return requirements.map(requirement =>
      evaluateSPFxApiPermissionRequirement(requirement, payload, {
        validateAudience: options?.validateAudience
      })
    );
  } catch (error) {
    return requirements.map(requirement =>
      classifySPFxApiPermissionError(requirement, error)
    );
  }
}

function groupRequirementsByResourceEndpoint(
  requirements: readonly SPFxApiPermissionRequirement[]
): Map<string, SPFxApiPermissionRequirement[]> {
  const groups = new Map<string, SPFxApiPermissionRequirement[]>();

  for (const requirement of requirements) {
    const group = groups.get(requirement.resourceEndpoint);

    if (group) {
      group.push(requirement);
    } else {
      groups.set(requirement.resourceEndpoint, [requirement]);
    }
  }

  return groups;
}

function getTokenWithTimeout(
  tokenProvider: SPFxAadTokenProviderLike,
  resourceEndpoint: string,
  useCachedToken: boolean,
  timeoutMs: number
): Promise<string> {
  let timeoutHandle: ReturnType<typeof setTimeout> | undefined;
  const tokenPromise = tokenProvider.getToken(resourceEndpoint, { useCachedToken });
  const timeoutPromise = new Promise<string>((_resolve, reject) => {
    timeoutHandle = setTimeout(() => reject(new Error('timeout')), timeoutMs);
  });

  return Promise.race([tokenPromise, timeoutPromise]).finally(() => {
    if (timeoutHandle) {
      clearTimeout(timeoutHandle);
    }
  });
}
