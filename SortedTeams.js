/** @OnlyCurrentDoc */

const sortedTeamsSheet = thisSpreadsheet.getSheetByName("Sorted Teams");
const sortedTeamsCheckboxesRange = sortedTeamsSheet.getRange("G1:G2");
const onlyIncludeBuiltCheckbox = sortedTeamsSheet.getRange("G1");
const onlyIncludeReleasedCheckbox = sortedTeamsSheet.getRange("G2");
const sortToTopCheckbox = sortedTeamsSheet.getRange("G8");
const sortedTeamsRange = sortedTeamsSheet.getRange("A2:D");
const refreshFormulasCheckbox = "G4";
const metaStrengthThresholdRange = sortedTeamsSheet.getRange("G5");

/**
 * Grabs the current sorted teams from the sheet, then returns a map of chacter to metadata (pun intended):
 * - numMetaTeams
 * - sortedTeamsStrengths (strengths of all possible teams for the character in descending order)
 * - strongestTeams (teams with the max team strength)
 */
function calculateCharacterMetaData() {
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
    characterMetaData.maxTeamStrengthTierScore = characterMetaData.sortedTeamsStrengths[Math.round(characterMetaData.sortedTeamsStrengths.length * topXPercentStrongestTeam)] / maxTeamStrength;
    const maxTierScore = Math.max(characterMetaData.numMetaTeamsTierScore, characterMetaData.maxTeamStrengthTierScore);
    const minTierScore = Math.min(characterMetaData.numMetaTeamsTierScore, characterMetaData.maxTeamStrengthTierScore);
    characterMetaData.tierScore = maxTierScore * tierScoreStrongerWeight + minTierScore * (1-tierScoreStrongerWeight);
    characterMetaData.tier = tierListBreakpoints.find(breakpoint => breakpoint[0] <= characterMetaData.tierScore)[1];
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