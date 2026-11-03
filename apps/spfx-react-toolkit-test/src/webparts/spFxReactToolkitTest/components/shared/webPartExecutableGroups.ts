import { executableGroups } from '../generated/jsdoc-examples/catalogue';
import { selectExecutableGroupsForHost } from './selectExecutableGroupsForHost';

/** One shared view for the actual WebPart initial selection, explorer and availability. */
export const webPartExecutableGroups = selectExecutableGroupsForHost(executableGroups, 'webPart');
