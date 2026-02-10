/** @OnlyCurrentDoc */

const sortedTeamsSheetName = "Sorted Teams";
const refreshFormulasCheckbox = "G4";
const sortedTeamsCheckboxesRangeString = "G1:G2";
const sortToTopCheckboxString = "G8";
const sortedTeamsRangeString = "A2:D";
const metaStrengthThresholdRangeString = "G5";

var _sortedTeamsSheet;
var _sortedTeamsCheckboxesRange;

function getSortedTeamsSheet() {
  if (!_sortedTeamsSheet) {
    _sortedTeamsSheet = getSpreadsheet().getSheetByName(sortedTeamsSheetName);
  }
  return _sortedTeamsSheet;
}

function getSortedTeamsCheckboxesRange() {
  if (!_sortedTeamsCheckboxesRange) {
    _sortedTeamsCheckboxesRange = getSortedTeamsSheet().getRange(sortedTeamsCheckboxesRangeString);
  }
  return _sortedTeamsCheckboxesRange;
}

/**
 * Grabs the current sorted teams from the sheet, then returns a map of chacter to metadata (pun intended):
 * - numMetaTeams
 * - sortedTeamsStrengths (strengths of all possible teams for the character in descending order)
 * - strongestTeams (teams with the max team strength)
 */
function calculateCharacterMetaData(tierListParams) {
  const sheet = getSortedTeamsSheet();
  const sortedTeamsRange = sheet.getRange(sortedTeamsRangeString);
  const metaStrengthThresholdRange = sheet.getRange(metaStrengthThresholdRangeString);

  const sortedTeams = sortedTeamsRange.getValues();
  const metaStrengthThreshold = metaStrengthThresholdRange.getValue();
  const charMetaDatas = new Map();
  var maxNumMetaTeams = 0;
  var maxTeamStrength = sortedTeams[0][sortedTeams[0].length - 1];
  outer: for (var r = 0; r < sortedTeams.length; r++) {
    const strength = sortedTeams[r][sortedTeams[r].length - 1];
    for (var c = 0; c < 3; c++) {
      const character = sortedTeams[r][c];
      if (character == null || character == "") {
        break outer;
      }
      var charMetaData = charMetaDatas.get(character);
      if (charMetaData === undefined) {
        charMetaData = {
          numMetaTeams : 0,
          sortedTeamsStrengths : [],
          strongestTeams : []
        }
        charMetaDatas.set(character, charMetaData);
      }

      if (strength >= metaStrengthThreshold) {
        charMetaData.numMetaTeams++;
        if (charMetaData.numMetaTeams > maxNumMetaTeams) {
          maxNumMetaTeams = charMetaData.numMetaTeams;
        }
      }
      charMetaData.sortedTeamsStrengths.push(strength);
      if (charMetaData.sortedTeamsStrengths.length < 3 || strength >= charMetaData.sortedTeamsStrengths[2]) {
        const strongestTeam = [];
        for (var c1 = 0; c1 < sortedTeams[r].length; c1++) {
          strongestTeam.push(sortedTeams[r][c1]);
        }
        charMetaData.strongestTeams.push(strongestTeam);
      }
    }
  }

  // Normalize numMetaTeams and maxTeamStrength as "tier score".
  charMetaDatas.forEach((characterMetaData, character, map) => {
    characterMetaData.numMetaTeamsTierScore = characterMetaData.numMetaTeams / maxNumMetaTeams;
    characterMetaData.maxTeamStrengthTierScore = characterMetaData.sortedTeamsStrengths[Math.round(characterMetaData.sortedTeamsStrengths.length * tierListParams.topXPercentStrongestTeam)] / maxTeamStrength;
    const maxTierScore = Math.max(characterMetaData.numMetaTeamsTierScore, characterMetaData.maxTeamStrengthTierScore);
    const minTierScore = Math.min(characterMetaData.numMetaTeamsTierScore, characterMetaData.maxTeamStrengthTierScore);
    characterMetaData.tierScore = maxTierScore * tierListParams.tierScoreStrongerWeight + minTierScore * (1-tierListParams.tierScoreStrongerWeight);
    characterMetaData.tier = tierListParams.tierListBreakpoints.find(breakpoint => breakpoint[0] <= characterMetaData.tierScore)[1];
    map.set(character, characterMetaData);
  });

  return charMetaDatas;
}

/**
 * Stores the existing values in sortedTeamsCheckboxesRange, which must be restored via:
 *  sortedTeamsCheckboxesRange.setValues(oldSortedTeamsCheckboxValues);
 * before the script ends.
 * 
 * Meanwhile, sets the given checkboxes to update the sorted teams.
 */
function setSortedTeamsCheckboxesAndGetOldValuesToRestoreLater(onlyIncludeBuilt, onlyIncludeReleased) {
  const sheet = getSortedTeamsSheet();
  const sortedTeamsCheckboxesRange = getSortedTeamsCheckboxesRange();
  const sortToTopCheckbox = sheet.getRange(sortToTopCheckboxString);

  const oldSortedTeamsCheckboxValues = sortedTeamsCheckboxesRange.getValues();
  // Deep copy.
  const tempSortedTeamsCheckboxValues = JSON.parse(JSON.stringify(oldSortedTeamsCheckboxValues));

  tempSortedTeamsCheckboxValues[0][0] = onlyIncludeBuilt;
  tempSortedTeamsCheckboxValues[1][0] = onlyIncludeReleased;
  sortedTeamsCheckboxesRange.setValues(tempSortedTeamsCheckboxValues);

  // We need to be sorted without any interference. Don't bother restoring this.
  sortToTopCheckbox.setValue(false);

  return oldSortedTeamsCheckboxValues;
}
