/**
 * Acquire one token per resource and return ordered delegated-scope diagnostics without React.
 *
 * @packageDocumentation
 * @docPage services/api-permission-precheck
 */
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
 * import { SPFxAadTokenProviderLike } from '@apvee/spfx-react-toolkit';
 *
 * function exampleUsage(spfxTokenProvider: SPFxAadTokenProviderLike) {
 *   const provider: SPFxAadTokenProviderLike = spfxTokenProvider;
 * }
 * ```
 */
export interface SPFxAadTokenProviderLike {
  /**
   * Acquires an access token; acquisition failures reject.
   * @param resourceEndpoint - Resource URL or application identifier.
   * @param options - Omitted follows the provider's native defaults; supplied cache preference and claims are forwarded.
   * @returns The access token.
   */
  getToken(
    resourceEndpoint: string,
    options?: {
      /** Whether the provider may reuse a cached token; omitted follows its native default. */
      readonly useCachedToken?: boolean;
      /** Claims challenge passed to the native provider; omitted requests no explicit claims. */
      readonly claims?: string
    }
  ): Promise<string>;
}

/**
 * Controls token acquisition and delegated-scope evaluation.
 *
 * @example
 * ```ts
 * import { SPFxApiPermissionPrecheckServiceOptions } from '@apvee/spfx-react-toolkit';
 *
 * const options: SPFxApiPermissionPrecheckServiceOptions = { useCachedToken: false, timeoutMs: 10000 };
 * ```
 */
export interface SPFxApiPermissionPrecheckServiceOptions {
  /**
   * Allows cached tokens unless false; omitted or undefined uses true.
   * @defaultValue true
   */
  readonly useCachedToken?: boolean;
  /**
   * Checks expected audiences unless false; omitted or undefined uses true.
   * @defaultValue true
   */
  readonly validateAudience?: boolean;
  /**
   * Acquisition timeout in milliseconds; omitted or undefined uses 15000. Timing out does not cancel the provider request.
   * @defaultValue 15000
   */
  readonly timeoutMs?: number;
  /**
   * Acquires each resource group in sequence when true; omitted/false acquires groups concurrently.
   * @defaultValue false
   */
  readonly sequentialResourceAcquisition?: boolean;
}

/**
 * Non-React delegated-scope diagnostics using one token acquisition per exact resource endpoint. Invalid requirements and group acquisition/evaluation failures become diagnostic results; malformed untyped configuration can still reject normalization. Empty config returns an empty result without acquisition. No React provider or lifecycle cleanup is required.
 *
 * @example
 * ```ts
 * import { SPFxAadTokenProviderLike, SPFxApiPermissionPrecheckService, createSPFxApiPermissionPrecheckService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(provider: SPFxAadTokenProviderLike) {
 *   const service: SPFxApiPermissionPrecheckService = createSPFxApiPermissionPrecheckService(provider);
 *   const results = await service.check({ graph: ['User.Read'] });
 * }
 * ```
 */
export interface SPFxApiPermissionPrecheckService {
  /**
   * Checks configured scopes without returning raw tokens.
   * @param config - Graph and custom API requirements.
   * @param options - Omitted enables cached tokens and audience checks, a 15000 ms timeout and concurrent resource groups.
   * @returns Invalid results first, then resource groups in first-seen order and scopes in group order; concurrent completion does not reorder them. Results contain no raw token and expose packageSolutionEntry for remediation.
   */
  check(
    config: SPFxApiPermissionPrecheckConfig,
    options?: SPFxApiPermissionPrecheckServiceOptions
  ): Promise<readonly SPFxApiPermissionCheckResult[]>;
}

/**
 * Creates permission diagnostics without React state or provider context. The caller's token provider controls interaction: this service installs no passive popup/redirect handlers, requests no claims challenge and adds no retry. Timeout clears its timer but does not cancel token acquisition. Available is diagnostic; the target API still authorizes requests.
 * @param tokenProvider - Initialized token provider owned by the caller.
 * @returns A service that evaluates delegated scopes from token payloads.
 *
 * @example Basic usage
 * ```ts
 * import { SPFxAadTokenProviderLike, createSPFxApiPermissionPrecheckService } from '@apvee/spfx-react-toolkit';
 *
 * async function exampleUsage(provider: SPFxAadTokenProviderLike) {
 *   const service = createSPFxApiPermissionPrecheckService(provider);
 *   const results = await service.check({ graph: ['User.Read'] });
 * }
 * ```
 * @useWhen You need delegated-scope diagnostics outside React using a caller-supplied token provider.
 *
 *
 * @example Check real delegated token diagnostics
 * ```ts
 * import type { SPFxAadTokenProviderLike } from '@apvee/spfx-react-toolkit';
 *
 * import { createSPFxApiPermissionPrecheckService } from '@apvee/spfx-react-toolkit';
 * async function demonstrate(tokenProvider: SPFxAadTokenProviderLike) {
 *   const diagnostics = createSPFxApiPermissionPrecheckService(tokenProvider);
 *   const results = await diagnostics.check({
 *     graph: ['User.Read', { scope: 'Mail.Read', required: false }],
 *   }, { useCachedToken: false, timeoutMs: 10000 });
 *   for (const result of results) console.log(result.scope, result.status, result.packageSolutionEntry);
 * }
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
