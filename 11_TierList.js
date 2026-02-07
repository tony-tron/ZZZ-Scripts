/** @OnlyCurrentDoc */

const tierListSheetName = "Tier List";
const tierListSheet = thisSpreadsheet.getSheetByName(tierListSheetName);
const recalculateTierListCheckbox = "I3";
const tierListBreakpointsRange = tierListSheet.getRange("F2:G");
const tierListBreakpoints = tierListBreakpointsRange.getValues();
const characterOutputRange = tierListSheet.getRange("A2:E");
const tierListOnlyIncludeBuiltCheckbox = tierListSheet.getRange("I6");
const tierListOnlyIncludeReleasedCheckbox = tierListSheet.getRange("I7");
const tierScoreStrongerWeight = tierListSheet.getRange("I8").getValue();
const topXPercentStrongestTeam = tierListSheet.getRange("I9").getValue();

function updateTierListSheet() {
  const dataValues = tierListSheet.getDataRange().getValues();
  const tempOnlyIncludeBuilt = dataValues
    [tierListOnlyIncludeBuiltCheckbox.getRow() - 1]
    [tierListOnlyIncludeBuiltCheckbox.getColumn() - 1];
  const tempOnlyIncludeReleased = dataValues
    [tierListOnlyIncludeReleasedCheckbox.getRow() - 1]
    [tierListOnlyIncludeReleasedCheckbox.getColumn() - 1];
  const oldSortedTeamsCheckboxValues = setSortedTeamsCheckboxesAndGetOldValuesToRestoreLater(
    tempOnlyIncludeBuilt, tempOnlyIncludeReleased);

  const charMetaDatas = calculateCharacterMetaData();
  const characterOutputs = [];
  for (const character of getCharacterNames()) {
    const charMetaData = charMetaDatas.get(character);
    if (charMetaData === undefined) continue;
    characterOutputs.push([
      character,
      charMetaData.numMetaTeamsTierScore,
      charMetaData.maxTeamStrengthTierScore,
      charMetaData.tierScore,
      charMetaData.tier
    ]);
  }
  characterOutputs.sort((output1, output2) => output2[3] - output1[3]) // Sort based on tierScore, index 3.

  for (var i = characterOutputs.length; i < characterOutputRange.getNumRows(); i++) {
    characterOutputs.push([null, null, null, null, null]);
  }
  characterOutputRange.clearContent().setValues(characterOutputs);

  updateTierListFormatting();

  sortedTeamsCheckboxesRange.setValues(oldSortedTeamsCheckboxValues);
}

/** Adds borders between tiers. (Colors are handled via Conditional Formatting) */
function updateTierListFormatting() {
  characterOutputRange.setBorder(false, false, false, false, false, false);
  const output = characterOutputRange.getValues();

  const tierRanges = [];
  var currTier = "Tier 0";
  var currTierStartRow = 0;
  for (var r = 0; r < output.length; r++) {
    const tier = output[r][output[r].length - 1];
    if (tier !== currTier) {
      tierRanges.push(
        tierListSheet.getRange(characterOutputRange.getRow() + currTierStartRow,
        characterOutputRange.getColumn(),
        r - currTierStartRow,
        characterOutputRange.getNumColumns()
        ));
      currTier = tier;
      currTierStartRow = r;
    }
    if (tier == null || tier === "") {
      break;
    }
  }
  
  for (var tier = 0; tier < tierRanges.length; tier++) {
    tierRanges[tier].setBorder(tier > 0, false, tier < tierRanges.length - 1, false, false, false);
  }
}
