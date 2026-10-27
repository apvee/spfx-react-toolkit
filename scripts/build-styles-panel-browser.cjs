const fs = require('node:fs');
const path = require('node:path');
const webpack = require('webpack');
const root = path.resolve(__dirname, '..');
const fixture = path.join(root, 'tests/fixtures/styles-panel-browser');
const output = path.join(root, 'temp/styles-panel-browser');
const cardsOnly = process.env.STYLES_PANEL_CARDS_ONLY === '1';
if (cardsOnly && ['DESCRIPTION_LIST', 'PAGINATION', 'STAGE', 'SHELL', 'DATA_GRID', 'FORM', 'GROUPS'].some(key => process.env[`STYLES_PANEL_${key}_ONLY`] === '1')) throw new Error('Cards-only fixture cannot be combined with another Only flag.');
const descriptionListOnly = process.env.STYLES_PANEL_DESCRIPTION_LIST_ONLY === '1';
if (descriptionListOnly && ['PAGINATION', 'STAGE', 'SHELL', 'DATA_GRID', 'FORM', 'GROUPS'].some(key => process.env[`STYLES_PANEL_${key}_ONLY`] === '1')) throw new Error('DescriptionList-only fixture cannot be combined with another Only flag.');
const paginationOnly = process.env.STYLES_PANEL_PAGINATION_ONLY === '1';
if (paginationOnly && ['STAGE', 'SHELL', 'DATA_GRID', 'FORM', 'GROUPS'].some(key => process.env[`STYLES_PANEL_${key}_ONLY`] === '1')) throw new Error('Pagination-only fixture cannot be combined with another Only flag.');
const stageOnly = process.env.STYLES_PANEL_STAGE_ONLY === '1';
const shellOnly = process.env.STYLES_PANEL_SHELL_ONLY === '1';
const gridOnly = process.env.STYLES_PANEL_DATA_GRID_ONLY === '1';
const formOnly = process.env.STYLES_PANEL_FORM_ONLY === '1';
if (shellOnly && gridOnly) throw new Error('Shell-only and Grid-only fixture flags are mutually exclusive.');
if (stageOnly && [shellOnly, gridOnly, process.env.STYLES_PANEL_FORM_ONLY === '1', process.env.STYLES_PANEL_GROUPS_ONLY === '1'].some(Boolean)) throw new Error('Stage-only fixture cannot be combined with another Only flag.');
const chartProbe = process.env.DATA_GRID_CHART_PROBE === '1';
if (chartProbe && !gridOnly) throw new Error('DATA_GRID_CHART_PROBE=1 requires STYLES_PANEL_DATA_GRID_ONLY=1 to select the isolated Grid fixture.');
webpack({ mode: 'development', devtool: 'source-map', entry: path.join(fixture, cardsOnly ? 'cards.tsx' : descriptionListOnly ? 'description-list.tsx' : paginationOnly ? 'pagination.tsx' : stageOnly ? 'stage.tsx' : shellOnly ? 'web-part-shell.tsx' : gridOnly ? 'data-grid.tsx' : formOnly ? 'form.tsx' : 'index.tsx'),
  output: { path: output, filename: 'fixture.js' },
  resolve: { extensions: ['.tsx', '.ts', '.js'], alias: {
    ...require('./localization-webpack.internal.cjs').nativeResourceAliases({ packageRoot: path.join(root, 'packages/spfx-react-toolkit') }),
    '@apvee/spfx-react-toolkit$': path.join(fixture, cardsOnly ? 'cards-toolkit-boundary.ts' : descriptionListOnly ? 'description-list-toolkit-boundary.ts' : paginationOnly ? 'pagination-toolkit-boundary.ts' : stageOnly ? 'stage-toolkit-boundary.ts' : shellOnly ? 'web-part-shell-toolkit-boundary.ts' : gridOnly ? 'data-grid-toolkit-boundary.ts' : formOnly ? 'form-toolkit-boundary.ts' : 'toolkit-boundary.ts'),
    ...(!cardsOnly && !descriptionListOnly && !paginationOnly && !stageOnly && !shellOnly && !gridOnly && !formOnly ? { '@apvee/spfx-react-toolkit/hooks$': path.join(fixture, 'toolkit-boundary.ts'), '@apvee/spfx-react-toolkit/styles$': path.join(fixture, 'toolkit-boundary.ts') } : {}),
    ...(shellOnly ? { '@microsoft/sp-core-library$': path.join(fixture, '../web-part-shell-browser/spfx-services.ts'), '@microsoft/sp-component-base$': path.join(fixture, '../web-part-shell-browser/spfx-services.ts') } : {}),
    ...(chartProbe ? { '@fluentui/react-charts$': path.join(root, 'tests/fixtures/sx-browser/data-grid-chart-probe.tsx'), '@data-grid/actual-charts$': path.join(root, 'node_modules/@fluentui/react-charts/lib/index.js') } : {}),
    // Resolve modern Fluent/Griffel ESM to the host's legacy React 17 entry.
    'react/jsx-runtime$': require.resolve('react/jsx-runtime.js'),
  } },
  plugins: shellOnly || stageOnly ? [] : cardsOnly ? [] : descriptionListOnly ? [] : paginationOnly ? [] : formOnly ? [] : [new webpack.NormalModuleReplacementPlugin(/^\.\/useSPFx(Teams|ThemeInfo)$/, resource => {
    if (resource.context === path.join(root, 'packages/spfx-react-toolkit/lib/hooks')) resource.request = path.join(fixture, 'spfx-boundaries.ts');
  })],
  module: { rules: [{ test: /\.tsx?$/, use: path.join(__dirname, 'sx-browser-typescript-loader.cjs') }] }
}, (error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString({ all: false, errors: true })); process.exitCode = 1; return; }
  fs.copyFileSync(path.join(fixture, cardsOnly ? '../cards-browser/index.html' : descriptionListOnly ? '../description-list-browser/index.html' : paginationOnly ? '../pagination-browser/index.html' : shellOnly ? '../web-part-shell-browser/index.html' : 'index.html'), path.join(output, 'index.html'));
  console.log(stats.toString({ all: false, timings: true, assets: true, warnings: true }));
  console.log(`Actual StylesPanel fixture: ${output}`);
});
