// Isolated native app. Real installed SDK declarations provide the host contracts;
// test-only consumers are deliberately outside the native root graph.
module.exports = {
  'apps/spfx-react-toolkit-test/tsconfig.json': JSON.stringify({ compilerOptions: { target: 'ES2020', module: 'ESNext', moduleResolution: 'node', jsx: 'react', skipLibCheck: true, strict: true, rootDir: 'src', outDir: 'lib' }, include: ['src/**/*'] }),
  'apps/spfx-react-toolkit-test/config/config.json': JSON.stringify({ bundles: {
    web: { components: [{ entrypoint: './lib/Web.js', manifest: './src/Web.manifest.json' }] },
    top: { components: [{ entrypoint: './lib/Top.js', manifest: './src/Top.manifest.json' }] }
  }, localizedResources: {} }),
  'apps/spfx-react-toolkit-test/src/Web.manifest.json': JSON.stringify({ id: '11111111-1111-4111-8111-111111111111', componentType: 'WebPart', supportedHosts: ['SharePointWebPart'], preconfiguredEntries: [{ properties: { message: 'hello' } }] }),
  'apps/spfx-react-toolkit-test/src/Top.manifest.json': JSON.stringify({ id: '22222222-2222-4222-8222-222222222222', componentType: 'Extension', extensionType: 'ApplicationCustomizer' }),
  'apps/spfx-react-toolkit-test/config/serve.json': JSON.stringify({ serveConfigurations: { top: { customActions: { '22222222-2222-4222-8222-222222222222': { location: 'ClientSideExtension.ApplicationCustomizer' } } } } }),
  'apps/spfx-react-toolkit-test/config/package-solution.json': JSON.stringify({ solution: { features: [{ assets: { elementManifests: ['elements.xml'] } }] } }),
  'apps/spfx-react-toolkit-test/sharepoint/assets/elements.xml': '<Elements><CustomAction Location="ClientSideExtension.ApplicationCustomizer" ClientSideComponentId="22222222-2222-4222-8222-222222222222" /></Elements>',
  'apps/spfx-react-toolkit-test/src/Web.ts': `import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { mount } from './host';
export default class Web extends BaseClientSideWebPart<{ message: string }> {
  public render(): void { mount(); }
}
`,
  'apps/spfx-react-toolkit-test/src/Top.ts': `import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
export default class Top extends BaseApplicationCustomizer<{}> {
  public onInit(): Promise<void> { return Promise.resolve(); }
}
`,
  'apps/spfx-react-toolkit-test/src/catalogue.ts': `export interface Entry { load(): Promise<{ Demo: () => number }>; }
export const entries: readonly Entry[] = [{ load: () => import('./scenario') }];
`,
  'apps/spfx-react-toolkit-test/src/host.ts': `import { entries, type Entry } from './catalogue';
export function mount(): void { entries.forEach((entry: Entry) => { entry.load().then(module => module.Demo()); }); }
`,
  'apps/spfx-react-toolkit-test/src/scenario.ts': 'export function Demo(): number { return 1; }\n'
};
