const fs = require('node:fs');
const path = require('node:path');
const { measureTreeShaking, baselineSnapshot } = require('./bundle-contract.internal.cjs');
function parseArgs(args) {
  const options = {mode:'enforce',output:'.docs/maintenance/evidence/tree-shaking/bundle-contract',attribution:false};
  for (let i = 0; i < args.length; i++) {
    if(args[i] === '--attribution') {options.attribution=true;continue;}
    if (!['--consumer-root', '--output', '--mode', '--contract', '--fixtures', '--write-baseline'].includes(args[i]) || !args[i + 1]) throw new Error(`Invalid argument ${args[i]}`);
    options[args[i].slice(2)] = args[++i];
  }
  if (!options['consumer-root']) throw new Error('Required: --consumer-root PATH');
  if (!['enforce','audit'].includes(options.mode)) throw new Error('mode must be audit or enforce');
  if (options['write-baseline'] && (options.mode !== 'enforce' || options.fixtures)) throw new Error('Baseline requires complete enforce mode');
  return options;
}
async function main(args) {
  const options=parseArgs(args);
  const contract = JSON.parse(fs.readFileSync(path.resolve(options.contract || path.join(__dirname, '../tests/fixtures/tree-shaking/contract.json'))));
  const input={ consumerRoot: options['consumer-root'], outputDir: options.output, mode: options.mode, contract, fixtureIds: options.fixtures?.split(',') };
  const report = await measureTreeShaking(input);
  console.log(`${report.mode}: passed=${report.passed}; ${report.failures.length} failures; ${report.warnings.length} warnings; ${report.durationMs}ms`);
  if (options['write-baseline']) fs.writeFileSync(path.resolve(options['write-baseline']),JSON.stringify(baselineSnapshot(report),null,2)+'\n');
  if (options.attribution) await measureTreeShaking({...input,mode:'audit',attribution:true,outputDir:path.join(options.output,'attribution')});
  if (report.mode === 'enforce' && !report.passed) process.exitCode = 1;
  return report;
}
if (require.main === module) main(process.argv.slice(2)).catch(error => { console.error(error); process.exitCode = 1; });
module.exports = { main, parseArgs };
