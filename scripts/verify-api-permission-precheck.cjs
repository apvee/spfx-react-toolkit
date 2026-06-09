const assert = require('assert');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const root = path.resolve(__dirname, '..');
const outDir = path.join(os.tmpdir(), 'spfx-api-permission-precheck-verify');

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

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
assert.strictEqual(optionalCustom[0].id, 'api:inventory:Inventory.Read');
assert.strictEqual(optionalCustom[0].required, false);
assert.strictEqual(optionalCustom[0].label, 'Inventory read');
assert.deepStrictEqual(optionalCustom[0].expectedAudiences, ['api://inventory']);

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

console.log('api permission precheck verification passed');
