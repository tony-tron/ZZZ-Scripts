/** @OnlyCurrentDoc */

const shiyuDefenseFrontier4Sheet = thisSpreadsheet.getSheetByName("Shiyu Defense - Frontier 4");
const minShiyuDefenseFrontier4TeamStrength = shiyuDefenseFrontier4Sheet.getRange("G4").getValue();
const maxShiyuDefenseFrontier4Options = shiyuDefenseFrontier4Sheet.getRange("G5").getValue();
const recalculateShiyuDefenseFrontier4Checkbox = "G2";

const shiyuDefenseFrontier4DistinctTeams = shiyuDefenseFrontier4Sheet.getRange("A2:D");
const shiyuDefenseFrontier4TeamsRow = shiyuDefenseFrontier4DistinctTeams.getRow();
const shiyuDefenseFrontier4TeamsColumn = shiyuDefenseFrontier4DistinctTeams.getColumn();

const shiyuDefenseFrontier4BuffsRange = shiyuDefenseFrontier4Sheet.getRange("F8:G");
var shiyuDefenseFrontier4Team1BuffExpressions = [];
var shiyuDefenseFrontier4Team2BuffExpressions = [];

function initalizeShiyuDefenseFrontier4BuffExpressions() {
  shiyuDefenseFrontier4Team1BuffExpressions = [];
  shiyuDefenseFrontier4Team2BuffExpressions = [];
  const buffNamesAndExpressions = shiyuDefenseFrontier4BuffsRange.getValues();

  var expression;
  var r = 0;
  // Team 1 buffs first.
  for (; r < buffNamesAndExpressions.length; r++) {
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    shiyuDefenseFrontier4Team1BuffExpressions.push(expression);
  }
  // Skip over empty cells.
  for (; r < buffNamesAndExpressions.length; r++) {
    name = buffNamesAndExpressions[r][0];
    expression = buffNamesAndExpressions[r][1];
    if ((name == null || name == "") && (expression == null || expression == "")) {
      continue;
    } else if (name != null && name != "") {
      // We reached the header for the next section.
      r++;
      break;
    }
  }
  // Team 2 buffs.
  for (; r < buffNamesAndExpressions.length; r++) {
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    shiyuDefenseFrontier4Team2BuffExpressions.push(expression);
  }
}

function updateShiyuDefenseFrontier4Sheet() {
  initalizeShiyuDefenseFrontier4BuffExpressions();
  clearShiyuDefenseFrontier4Teams();
  const allTeams = getAllTeams(minShiyuDefenseFrontier4TeamStrength);
  const teamPairs = computeBestDistinctTeamPairs(allTeams, shiyuDefenseFrontier4Team1BuffExpressions, shiyuDefenseFrontier4Team2BuffExpressions);
  const sortedPairs = teamPairs.sort((pair1, pair2) => pair2.minStrength() - pair1.minStrength() || pair2.totalStrength() - pair1.totalStrength());
  updateShiyuDefenseFrontier4DistinctTeamsSheet(sortedPairs);
}

function clearShiyuDefenseFrontier4Teams() {
  shiyuDefenseFrontier4DistinctTeams.clearContent().breakApart();
}

function updateShiyuDefenseFrontier4DistinctTeamsSheet(teamPairs) {
  if (teamPairs.length == 0) {
    shiyuDefenseFrontier4Sheet.getRange(shiyuDefenseFrontier4TeamsRow, shiyuDefenseFrontier4TeamsColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross()
  }
  var outputValues = [];
  var count = 0;
  for (var i = 0; i < teamPairs.length && i < maxShiyuDefenseFrontier4Options; i++) {
    var teamPair = teamPairs[i];
    var team1 = teamPair.team1;
    var team2 = teamPair.team2;
    var strengthString =
      team1.strength +  " + " + teamPair.team1Bonus + " \n+ " +
      team2.strength + " + " + teamPair.team2Bonus + "\n= " + teamPair.totalStrength() + " (min=" + teamPair.minStrength() + ")";

    outputValues.push([...team1.characters, strengthString]);
    outputValues.push([...team2.characters, ""]);
    outputValues.push(["", "", "", ""]);
    count++;
  }

  if (count > 0) {
    shiyuDefenseFrontier4Sheet.getRange(shiyuDefenseFrontier4TeamsRow, shiyuDefenseFrontier4TeamsColumn, count * 3, 4).setValues(outputValues);
    shiyuDefenseFrontier4Sheet.getRange(shiyuDefenseFrontier4TeamsRow, shiyuDefenseFrontier4TeamsColumn + 3, count * 3, 1).setVerticalAlignment('middle').setHorizontalAlignment('center');

    for (var i = 0; i < count; i++) {
      shiyuDefenseFrontier4Sheet.getRange(shiyuDefenseFrontier4TeamsRow + i * 3, shiyuDefenseFrontier4TeamsColumn + 3, 2, 1).mergeVertically();
    }
  }
}