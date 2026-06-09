const assert = require('assert');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const root = path.resolve(__dirname, '..');
const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spfx-api-permission-precheck-verify-'));

const sourceFiles = [
  'src/helpers/spfx-api-permission-precheck.helpers.ts'
];

if (fs.existsSync(path.join(root, 'src/services/spfx-api-permission-precheck.service.ts'))) {
  sourceFiles.push('src/services/spfx-api-permission-precheck.service.ts');
}

execFileSync(
  'npx',
  [
    'tsc',
    ...sourceFiles,
    '--module',
    'commonjs',
    '--target',
    'es2020',
    '--skipLibCheck',
    '--esModuleInterop',
    '--rootDir',
    'src',
    '--outDir',
    outDir,
    '--pretty',
    'false'
  ],
  { cwd: root, stdio: 'inherit' }
);

const helpers = require(path.join(outDir, 'helpers/spfx-api-permission-precheck.helpers.js'));
const servicePath = path.join(outDir, 'services/spfx-api-permission-precheck.service.js');
const service = fs.existsSync(servicePath) ? require(servicePath) : undefined;

function createJwt(payload) {
  const encoded = Buffer.from(JSON.stringify(payload))
    .toString('base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');

  return `header.${encoded}.signature`;
}

function createPayload(payload) {
  return helpers.decodeSPFxJwtPayload(createJwt(payload));
}

const now = Math.floor(Date.now() / 1000);

assert.deepStrictEqual(helpers.normalizeSPFxApiPermissionRequirements({}), []);

assert.strictEqual(
  helpers.summarizeSPFxApiPermissionResults([]).configurationState,
  'idle'
);
assert.strictEqual(
  helpers.summarizeSPFxApiPermissionResults([], true).configurationState,
  'checking'
);

const requirements = helpers.normalizeSPFxApiPermissionRequirements({
  graph: [' Sites.Read.All ', { scope: 'Files.Read.All', satisfies: ['Files.ReadWrite.All'] }],
  customApis: [
    {
      id: 'orders',
      name: 'Orders API',
      resource: 'api://orders',
      scopes: ['Orders.Read']
    }
  ]
});

assert.strictEqual(requirements.length, 3);
assert.strictEqual(requirements[0].id, 'graph:Sites.Read.All');
assert.strictEqual(requirements[0].resourceName, 'Microsoft Graph');
assert.strictEqual(requirements[0].resourceEndpoint, 'https://graph.microsoft.com');
assert.strictEqual(requirements[0].packageResource, 'Microsoft Graph');
assert.deepStrictEqual(requirements[0].expectedAudiences, [
  'https://graph.microsoft.com',
  '00000003-0000-0000-c000-000000000000'
]);
assert.strictEqual(requirements[0].required, true);
assert.strictEqual(requirements[0].scope, 'Sites.Read.All');
assert.strictEqual(requirements[0].adminMessage, 'Approve Microsoft Graph / Sites.Read.All in SharePoint Admin Center API access.');
assert.strictEqual(requirements[1].satisfies[0], 'Files.ReadWrite.All');
assert.strictEqual(requirements[2].id, 'api:orders:Orders.Read');
assert.strictEqual(requirements[2].resourceName, 'Orders API');
assert.strictEqual(requirements[2].resourceEndpoint, 'api://orders');
assert.strictEqual(requirements[2].packageResource, 'Orders API');

const optionalCustom = helpers.normalizeSPFxApiPermissionRequirements({
  customApis: [
    {
      name: 'Inventory API',
      resource: 'api://inventory',
      packageResource: 'Inventory Package',
      expectedAudiences: ['api://inventory'],
      scopes: [{ scope: ' Inventory.Read ', required: false, label: 'Inventory read' }]
    }
  ]
});
assert.strictEqual(optionalCustom[0].id, 'api:api://inventory:Inventory.Read');
assert.strictEqual(optionalCustom[0].required, false);
assert.strictEqual(optionalCustom[0].label, 'Inventory read');
assert.deepStrictEqual(optionalCustom[0].expectedAudiences, ['api://inventory']);
assert.strictEqual(
  helpers.createSPFxApiPermissionRequirementId('api://inventory', 'Inventory.Read', 'orders-api'),
  'api:orders-api:Inventory.Read'
);
assert.strictEqual(
  helpers.createSPFxApiPermissionRequirementId('api://inventory', 'Inventory.Read'),
  'api:api://inventory:Inventory.Read'
);

const deduped = helpers.normalizeSPFxApiPermissionRequirements({
  graph: ['User.Read', 'User.Read'],
  requirements: [
    {
      id: 'explicit-user-read',
      resourceName: 'Microsoft Graph',
      resourceEndpoint: 'https://graph.microsoft.com',
      packageResource: 'Microsoft Graph',
      scope: 'User.Read',
      kind: 'delegatedScope',
      required: true,
      adminMessage: 'explicit duplicate'
    },
    {
      id: 'explicit-mail-read',
      resourceName: 'Microsoft Graph',
      resourceEndpoint: 'https://graph.microsoft.com',
      packageResource: 'Microsoft Graph',
      scope: 'Mail.Read',
      kind: 'delegatedScope',
      required: true,
      adminMessage: 'explicit unique'
    }
  ]
});
assert.strictEqual(deduped.length, 2);
assert.strictEqual(deduped[0].id, 'graph:User.Read');
assert.strictEqual(deduped[1].id, 'explicit-mail-read');

const graphPayload = createPayload({
  aud: 'https://graph.microsoft.com',
  scp: 'Sites.Read.All Files.ReadWrite.All User.Read',
  tid: 'tenant-id',
  exp: now + 3600,
  nbf: now - 30
});
assert.ok(graphPayload);
assert.deepStrictEqual(
  helpers.extractSPFxDelegatedScopes(graphPayload),
  ['Sites.Read.All', 'Files.ReadWrite.All', 'User.Read']
);

const exactResult = helpers.evaluateSPFxApiPermissionRequirement(requirements[0], graphPayload);
assert.strictEqual(exactResult.status, 'available');
assert.strictEqual(exactResult.severity, 'ok');
assert.strictEqual(exactResult.matchedScope, 'Sites.Read.All');
assert.strictEqual(exactResult.audience, 'https://graph.microsoft.com');
assert.strictEqual(exactResult.tenantId, 'tenant-id');

const satisfiesResult = helpers.evaluateSPFxApiPermissionRequirement(requirements[1], graphPayload);
assert.strictEqual(satisfiesResult.status, 'available');
assert.strictEqual(satisfiesResult.matchedScope, 'Files.ReadWrite.All');

const mismatchPayload = createPayload({
  aud: 'https://graph.microsoft.com',
  scp: 'sites.read.all',
  exp: now + 3600
});
const caseResult = helpers.evaluateSPFxApiPermissionRequirement(requirements[0], mismatchPayload);
assert.strictEqual(caseResult.status, 'missingScope');
assert.strictEqual(caseResult.caseMismatchScope, 'sites.read.all');

const wrongAudiencePayload = createPayload({
  aud: 'api://wrong',
  scp: 'Sites.Read.All',
  exp: now + 3600
});
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], wrongAudiencePayload).status,
  'audienceMismatch'
);
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], wrongAudiencePayload, { validateAudience: false }).status,
  'available'
);

const rolesOnlyPayload = createPayload({
  aud: 'https://graph.microsoft.com',
  roles: ['Sites.Read.All'],
  exp: now + 3600
});
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], rolesOnlyPayload).status,
  'unsupportedPermissionKind'
);

assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], createPayload({
    aud: 'https://graph.microsoft.com',
    scp: 'Sites.Read.All',
    exp: now - 1
  })).status,
  'tokenExpired'
);
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], createPayload({
    aud: 'https://graph.microsoft.com',
    scp: 'Sites.Read.All',
    nbf: now + 3600
  })).status,
  'tokenExpired'
);

const missingResult = helpers.evaluateSPFxApiPermissionRequirement(
  requirements[0],
  createPayload({
    aud: 'https://graph.microsoft.com',
    scp: 'Mail.Read',
    exp: now + 3600
  })
);
assert.strictEqual(missingResult.status, 'missingScope');

const invalidRequirement = {
  ...requirements[0],
  scope: ''
};
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(invalidRequirement, graphPayload).status,
  'invalidRequirement'
);

const missingKindRequirement = {
  id: 'missing-kind',
  resourceName: 'Missing Kind API',
  resourceEndpoint: 'api://missing-kind',
  packageResource: 'Missing Kind API',
  scope: 'MissingKind.Read',
  required: true,
  adminMessage: 'Approve Missing Kind API / MissingKind.Read in SharePoint Admin Center API access.'
};
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(missingKindRequirement, graphPayload).status,
  'invalidRequirement'
);

assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], undefined).status,
  'unsupportedToken'
);

assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('AADSTS65001: consent required')).status,
  'consentRequired'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('login required with MFA')).status,
  'interactionRequired'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(
    requirements[0],
    new Error('SPFx API permission precheck passive mode blocked an authentication popup.')
  ).status,
  'interactionRequired'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('claims challenge received')).status,
  'claimsChallenge'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('conditional access policy blocked access')).status,
  'conditionalAccessBlocked'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('user assignment is required')).status,
  'accessRestricted'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('invalid resource not found')).status,
  'resourceMisconfigured'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('request timeout')).status,
  'timeout'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('network throttle temporarily unavailable')).status,
  'transientFailure'
);
assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('unexpected failure')).status,
  'unknown'
);

const summaryReady = helpers.summarizeSPFxApiPermissionResults([exactResult]);
assert.strictEqual(summaryReady.configurationState, 'ready');
assert.strictEqual(summaryReady.isConfigured, true);
assert.strictEqual(summaryReady.available.length, 1);

const summaryAction = helpers.summarizeSPFxApiPermissionResults([missingResult]);
assert.strictEqual(summaryAction.configurationState, 'actionRequired');
assert.strictEqual(summaryAction.isConfigured, false);
assert.strictEqual(summaryAction.missing.length, 1);

const optionalWarning = helpers.evaluateSPFxApiPermissionRequirement(
  optionalCustom[0],
  createPayload({
    aud: 'api://inventory',
    scp: 'Inventory.Write',
    exp: now + 3600
  })
);
const summaryWarning = helpers.summarizeSPFxApiPermissionResults([exactResult, optionalWarning]);
assert.strictEqual(summaryWarning.configurationState, 'ready');
assert.strictEqual(summaryWarning.isConfigured, true);
assert.strictEqual(summaryWarning.warnings.length, 1);

const summaryCannotDetermine = helpers.summarizeSPFxApiPermissionResults([
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('unexpected failure'))
]);
assert.strictEqual(summaryCannotDetermine.configurationState, 'cannotDetermine');
assert.strictEqual(summaryCannotDetermine.unknown.length, 1);

assert.strictEqual(helpers.decodeSPFxJwtPayload('not-a-jwt'), undefined);
assert.strictEqual(helpers.decodeSPFxJwtPayload('header.not-json.signature'), undefined);

async function verifyService() {
  assert.ok(service, 'compiled service module should be present');
  assert.strictEqual(typeof service.createSPFxApiPermissionPrecheckService, 'function');

  const graphToken = createJwt({
    aud: 'https://graph.microsoft.com',
    scp: 'Sites.Read.All User.Read',
    exp: now + 3600
  });
  const customToken = createJwt({
    aud: 'api://orders',
    scp: 'Orders.Read',
    exp: now + 3600
  });

  const graphCalls = [];
  const graphOnlyPrecheck = service.createSPFxApiPermissionPrecheckService({
    getToken: async (resourceEndpoint, options) => {
      graphCalls.push({ resourceEndpoint, options });
      return graphToken;
    }
  });
  const graphOnlyResults = await graphOnlyPrecheck.check({
    graph: ['Sites.Read.All', 'User.Read']
  });
  assert.strictEqual(graphCalls.length, 1);
  assert.strictEqual(graphCalls[0].resourceEndpoint, 'https://graph.microsoft.com');
  assert.deepStrictEqual(graphCalls[0].options, { useCachedToken: true });
  assert.deepStrictEqual(graphOnlyResults.map(result => result.status), ['available', 'available']);

  const multiResourceCalls = [];
  const multiResourcePrecheck = service.createSPFxApiPermissionPrecheckService({
    getToken: async (resourceEndpoint, options) => {
      multiResourceCalls.push({ resourceEndpoint, options });
      return resourceEndpoint === 'https://graph.microsoft.com' ? graphToken : customToken;
    }
  });
  const multiResourceResults = await multiResourcePrecheck.check({
    graph: ['Sites.Read.All'],
    customApis: [
      {
        id: 'orders',
        name: 'Orders API',
        resource: 'api://orders',
        expectedAudiences: ['api://orders'],
        scopes: ['Orders.Read']
      }
    ]
  });
  assert.strictEqual(multiResourceCalls.length, 2);
  assert.deepStrictEqual(
    multiResourceCalls.map(call => call.resourceEndpoint).sort(),
    ['api://orders', 'https://graph.microsoft.com']
  );
  assert.deepStrictEqual(multiResourceResults.map(result => result.status), ['available', 'available']);

  let resolveSequentialGraphToken;
  const sequentialCalls = [];
  const sequentialPrecheck = service.createSPFxApiPermissionPrecheckService({
    getToken: async resourceEndpoint => {
      sequentialCalls.push(resourceEndpoint);

      if (resourceEndpoint === 'https://graph.microsoft.com') {
        return new Promise(resolve => {
          resolveSequentialGraphToken = () => resolve(graphToken);
        });
      }

      return customToken;
    }
  });
  const sequentialCheckPromise = sequentialPrecheck.check(
    {
      graph: ['Sites.Read.All'],
      customApis: [
        {
          id: 'orders',
          name: 'Orders API',
          resource: 'api://orders',
          expectedAudiences: ['api://orders'],
          scopes: ['Orders.Read']
        }
      ]
    },
    { sequentialResourceAcquisition: true }
  );
  await Promise.resolve();
  assert.deepStrictEqual(sequentialCalls, ['https://graph.microsoft.com']);
  assert.strictEqual(typeof resolveSequentialGraphToken, 'function');
  resolveSequentialGraphToken();
  const sequentialResults = await sequentialCheckPromise;
  assert.deepStrictEqual(sequentialCalls, ['https://graph.microsoft.com', 'api://orders']);
  assert.deepStrictEqual(sequentialResults.map(result => result.status), ['available', 'available']);

  const consentPrecheck = service.createSPFxApiPermissionPrecheckService({
    getToken: async () => {
      throw new Error('AADSTS65001: The user or administrator has not consented.');
    }
  });
  const consentResults = await consentPrecheck.check({
    graph: ['Mail.Read']
  });
  assert.strictEqual(consentResults.length, 1);
  assert.strictEqual(consentResults[0].status, 'consentRequired');

  const timeoutPrecheck = service.createSPFxApiPermissionPrecheckService({
    getToken: () => new Promise(resolve => setTimeout(() => resolve(graphToken), 50))
  });
  const timeoutResults = await timeoutPrecheck.check(
    { graph: ['Sites.Read.All'] },
    { timeoutMs: 1 }
  );
  assert.strictEqual(timeoutResults.length, 1);
  assert.strictEqual(timeoutResults[0].status, 'timeout');

  const invalidCalls = [];
  const invalidPrecheck = service.createSPFxApiPermissionPrecheckService({
    getToken: async resourceEndpoint => {
      invalidCalls.push(resourceEndpoint);
      return graphToken;
    }
  });
  const invalidResults = await invalidPrecheck.check({
    requirements: [
      {
        id: 'invalid',
        resourceName: 'Invalid API',
        resourceEndpoint: 'api://invalid',
        packageResource: 'Invalid API',
        scope: '',
        kind: 'delegatedScope',
        required: true,
        adminMessage: 'Invalid requirement'
      }
    ]
  });
  assert.strictEqual(invalidCalls.length, 0);
  assert.strictEqual(invalidResults.length, 1);
  assert.strictEqual(invalidResults[0].status, 'invalidRequirement');
}

verifyService()
  .then(() => {
    console.log('api permission precheck verification passed');
    fs.rmSync(outDir, { recursive: true, force: true });
  })
  .catch(error => {
    fs.rmSync(outDir, { recursive: true, force: true });
    throw error;
  });
