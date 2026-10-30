import { EntraGroupTrim } from '__TOOLKIT__';
import { SPFxWebPartProvider } from '__PROVIDER__';
import { runTrimProbe } from './access-trim-probe';
export async function run(host) { return runTrimProbe(EntraGroupTrim, SPFxWebPartProvider, 'entra-group', host); }
