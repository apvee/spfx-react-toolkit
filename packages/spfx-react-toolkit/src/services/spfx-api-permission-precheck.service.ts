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

export interface SPFxAadTokenProviderLike {
  getToken(
    resourceEndpoint: string,
    options?: { readonly useCachedToken?: boolean; readonly claims?: string }
  ): Promise<string>;
}

export interface SPFxApiPermissionPrecheckServiceOptions {
  readonly useCachedToken?: boolean;
  readonly validateAudience?: boolean;
  readonly timeoutMs?: number;
  readonly sequentialResourceAcquisition?: boolean;
}

export interface SPFxApiPermissionPrecheckService {
  check(
    config: SPFxApiPermissionPrecheckConfig,
    options?: SPFxApiPermissionPrecheckServiceOptions
  ): Promise<readonly SPFxApiPermissionCheckResult[]>;
}

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
