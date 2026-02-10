/** @OnlyCurrentDoc */

const tierListSheetName = "Tier List";
const recalculateTierListCheckbox = "I3";
const tierListBreakpointsRange = "F2:G";
const characterOutputRange = "A2:E";
const tierListOnlyIncludeBuiltCheckbox = "I6";
const tierListOnlyIncludeReleasedCheckbox = "I7";
const tierScoreStrongerWeightRange = "I8";
const topXPercentStrongestTeamRange = "I9";

var _tierListSheet;

function getTierListSheet() {
  if (!_tierListSheet) {
    _tierListSheet = getSpreadsheet().getSheetByName(tierListSheetName);
  }
  return _tierListSheet;
}

function getTierListParams() {
  const sheet = getTierListSheet();
  return {
    tierScoreStrongerWeight: sheet.getRange(tierScoreStrongerWeightRange).getValue(),
    topXPercentStrongestTeam: sheet.getRange(topXPercentStrongestTeamRange).getValue(),
    tierListBreakpoints: sheet.getRange(tierListBreakpointsRange).getValues(),
  };
}

function updateTierListSheet() {
  const sheet = getTierListSheet();
  const tierListOnlyIncludeBuiltCheckboxRange = sheet.getRange(tierListOnlyIncludeBuiltCheckbox);
  const tierListOnlyIncludeReleasedCheckboxRange = sheet.getRange(tierListOnlyIncludeReleasedCheckbox);
  const characterOutputRangeObj = sheet.getRange(characterOutputRange);

  const dataValues = sheet.getDataRange().getValues();
  const tempOnlyIncludeBuilt = dataValues
    [tierListOnlyIncludeBuiltCheckboxRange.getRow() - 1]
    [tierListOnlyIncludeBuiltCheckboxRange.getColumn() - 1];
  const tempOnlyIncludeReleased = dataValues
    [tierListOnlyIncludeReleasedCheckboxRange.getRow() - 1]
    [tierListOnlyIncludeReleasedCheckboxRange.getColumn() - 1];
  const oldSortedTeamsCheckboxValues = setSortedTeamsCheckboxesAndGetOldValuesToRestoreLater(
    tempOnlyIncludeBuilt, tempOnlyIncludeReleased);

  const charMetaDatas = calculateCharacterMetaData(getTierListParams());
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

  for (var i = characterOutputs.length; i < characterOutputRangeObj.getNumRows(); i++) {
    characterOutputs.push([null, null, null, null, null]);
  }
  characterOutputRangeObj.clearContent().setValues(characterOutputs);

  updateTierListFormatting(characterOutputRangeObj);

  getSortedTeamsCheckboxesRange().setValues(oldSortedTeamsCheckboxValues);
}

/** Adds borders between tiers. (Colors are handled via Conditional Formatting) */
function updateTierListFormatting(characterOutputRange) {
  characterOutputRange.setBorder(false, false, false, false, false, false);
  const output = characterOutputRange.getValues();

  const tierRanges = [];
  var currTier = "Tier 0";
  var currTierStartRow = 0;
  const sheet = getTierListSheet();
  for (var r = 0; r < output.length; r++) {
    const tier = output[r][output[r].length - 1];
    if (tier !== currTier) {
      tierRanges.push(
        sheet.getRange(characterOutputRange.getRow() + currTierStartRow,
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
