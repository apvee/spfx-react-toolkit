import { SharePointGroupTrim } from '__TOOLKIT__';
import { SPFxWebPartProvider } from '__PROVIDER__';
import { runTrimProbe } from './access-trim-probe';
export async function run(host) { return runTrimProbe(SharePointGroupTrim, SPFxWebPartProvider, 'sharepoint-group', host); }
