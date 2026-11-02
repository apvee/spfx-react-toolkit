// Independent browser acceptance expectations, never a generator/navigation input.
const pilotTitles = {
  WebPartShell: ['Configure the actual web part Shell', 'Retain content through Shell presentation changes', 'Observe content remount between subtle and Card'],
  Pagination: ['Browse controlled rows', 'Browse controlled rows with pending completion', 'Show an empty controlled result', 'Browse orders across pagination modes'],
  useStableCallback: ['Report current count through a retained callback', 'Retain callback identity through rerenders'],
  useSx: ['Compare root domain and historical imports'],
  useSPFxPnPList: ['Read a user-selected list with observable results'],
};
function filterPilotGroups(groups) {
  const features = groups.flatMap(group => group.features);
  for (const [title, scenarios] of Object.entries(pilotTitles)) {
    const matches = features.filter(feature => feature.title === title);
    if (matches.length !== 1) throw new Error(`Missing pilot feature or duplicate: ${title}`);
    for (const scenario of scenarios) if (matches[0].scenarios.filter(entry => entry.title === scenario).length !== 1)
      throw new Error(`Missing pilot scenario or duplicate: ${title}/${scenario}`);
  }
  return groups.map(group => ({ ...group, features: group.features.filter(feature => Object.hasOwn(pilotTitles, feature.title))
    .map(feature => ({ ...feature, scenarios: feature.scenarios.filter(entry => pilotTitles[feature.title].includes(entry.title)) })) }))
    .filter(group => group.features.length > 0);
}
module.exports = { pilotTitles, filterPilotGroups };
