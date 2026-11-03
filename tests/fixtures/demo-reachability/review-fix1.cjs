// Independent counterexamples from the Task9 review; no checkout/evidence input.
const files = require('./files.cjs');
const app = 'apps/spfx-react-toolkit-test/';
const extensionId = '22222222-2222-4222-8222-222222222222';
const wrongId = '33333333-3333-4333-8333-333333333333';
const location = 'ClientSideExtension.ApplicationCustomizer';

function namespaceConsumer(statement, prelude = '') {
  return {
    [app + 'src/host.ts']: "import * as namespace from './scenario';\n" +
      'declare function leak(value: unknown): void;\n' +
      (prelude ? prelude + '\n' : '') +
      files[app + 'src/host.ts'].replace('entries.forEach(', statement + '; entries.forEach('),
  };
}
function serveWithoutAction({ removeProperty = false } = {}) {
  const serve = JSON.parse(files[app + 'config/serve.json']);
  if (removeProperty) delete serve.serveConfigurations.top.customActions;
  else delete serve.serveConfigurations.top.customActions[extensionId];
  return { [app + 'config/serve.json']: JSON.stringify(serve) };
}
function solutionWithoutAssociation({ removeAssets = false } = {}) {
  const solution = JSON.parse(files[app + 'config/package-solution.json']);
  if (removeAssets) delete solution.solution.features[0].assets;
  else solution.solution.features[0].assets.elementManifests = [];
  return { [app + 'config/package-solution.json']: JSON.stringify(solution) };
}
function nativePair() {
  const solution = JSON.parse(files[app + 'config/package-solution.json']);
  solution.solution.features[0].assets.elementManifests.push('instance.xml');
  return {
    [app + 'config/package-solution.json']: JSON.stringify(solution),
    [app + 'sharepoint/assets/instance.xml']: `<sp:Elements xmlns:sp='http://schemas.microsoft.com/sharepoint/'>\n` +
      `<!-- Wrong commented registration ${wrongId} -->\n` +
      `<sp:ClientSideComponentInstance Location='${location}' ComponentId='${extensionId}' />\n</sp:Elements>`,
  };
}
module.exports = { namespaceConsumer, serveWithoutAction, solutionWithoutAssociation, nativePair, extensionId, wrongId, location };
