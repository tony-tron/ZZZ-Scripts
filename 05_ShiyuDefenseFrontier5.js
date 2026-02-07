/** @OnlyCurrentDoc */

const shiyuDefenseFrontier5SheetName = "Shiyu Defense - Frontier 5";
const shiyuDefenseFrontier5Sheet = thisSpreadsheet.getSheetByName(shiyuDefenseFrontier5SheetName);
const minShiyuDefenseFrontier5TeamStrength = shiyuDefenseFrontier5Sheet.getRange("G4").getValue();
const maxShiyuDefenseFrontier5Options = shiyuDefenseFrontier5Sheet.getRange("G5").getValue();
const recalculateShiyuDefenseFrontier5Checkbox = "G2";

const shiyuDefenseFrontier5DistinctTeams = shiyuDefenseFrontier5Sheet.getRange("A2:D");
const shiyuDefenseFrontier5TeamsRow = shiyuDefenseFrontier5DistinctTeams.getRow();
const shiyuDefenseFrontier5TeamsColumn = shiyuDefenseFrontier5DistinctTeams.getColumn();

const shiyuDefenseFrontier5BuffsRange = shiyuDefenseFrontier5Sheet.getRange("F8:G");
var shiyuDefenseFrontier5Team1BuffExpressions = [];
var shiyuDefenseFrontier5Team2BuffExpressions = [];
var shiyuDefenseFrontier5Team3BuffExpressions = [];

function initalizeShiyuDefenseFrontier5BuffExpressions() {
  shiyuDefenseFrontier5Team1BuffExpressions = [];
  shiyuDefenseFrontier5Team2BuffExpressions = [];
  shiyuDefenseFrontier5Team3BuffExpressions = [];
  const buffNamesAndExpressions = shiyuDefenseFrontier5BuffsRange.getValues();

  var expression;
  var r = 0;
  // Team 1 buffs first.
  for (; r < buffNamesAndExpressions.length; r++) {
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    shiyuDefenseFrontier5Team1BuffExpressions.push(expression);
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
    shiyuDefenseFrontier5Team2BuffExpressions.push(expression);
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
  // Team 3 buffs.
  for (; r < buffNamesAndExpressions.length; r++) {
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    shiyuDefenseFrontier5Team3BuffExpressions.push(expression);
  }
}

function updateShiyuDefenseFrontier5Sheet() {
  initalizeShiyuDefenseFrontier5BuffExpressions();
  clearShiyuDefenseFrontier5Teams();
  const allTeams = getAllTeams(minShiyuDefenseFrontier5TeamStrength);
  const teamTriples = computeBestDistinctTeamTriples(allTeams, shiyuDefenseFrontier5Team1BuffExpressions, shiyuDefenseFrontier5Team2BuffExpressions, shiyuDefenseFrontier5Team3BuffExpressions);
  const sortedTriples = teamTriples.sort((triple1, triple2) => triple2.minStrength() - triple1.minStrength() || triple2.totalStrength() - triple1.totalStrength());
  updateShiyuDefenseFrontier5DistinctTeamsSheet(sortedTriples);
}

function clearShiyuDefenseFrontier5Teams() {
  shiyuDefenseFrontier5DistinctTeams.clearContent().breakApart();
}

function updateShiyuDefenseFrontier5DistinctTeamsSheet(teamTriples) {
  if (teamTriples.length == 0) {
    shiyuDefenseFrontier5Sheet.getRange(shiyuDefenseFrontier5TeamsRow, shiyuDefenseFrontier5TeamsColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross()
  }
  for (var i = 0; i < teamTriples.length && i < maxShiyuDefenseFrontier5Options; i++) {
    var teamTriple = teamTriples[i];
    var team1 = teamTriple.team1;
    var team2 = teamTriple.team2;
    var team3 = teamTriple.team3;
    var strengthString =
      team1.strength +  " + " + teamTriple.team1Bonus + " \n+ " +
      team2.strength + " + " + teamTriple.team2Bonus + "\n= " +
      team3.strength + " + " + teamTriple.team3Bonus + "\n= " + teamTriple.totalStrength() + " (min=" + teamTriple.minStrength() + ")";
    shiyuDefenseFrontier5Sheet.getRange(shiyuDefenseFrontier5TeamsRow + i * 4, shiyuDefenseFrontier5TeamsColumn, 1, 3).setValues([team1.characters]);
    shiyuDefenseFrontier5Sheet.getRange(shiyuDefenseFrontier5TeamsRow + 1 + i * 4, shiyuDefenseFrontier5TeamsColumn, 1, 3).setValues([team2.characters]);
    shiyuDefenseFrontier5Sheet.getRange(shiyuDefenseFrontier5TeamsRow + 2 + i * 4, shiyuDefenseFrontier5TeamsColumn, 1, 3).setValues([team3.characters]);
    shiyuDefenseFrontier5Sheet.getRange(shiyuDefenseFrontier5TeamsRow + i * 4, shiyuDefenseFrontier5TeamsColumn + 3, 3, 1).setValue(strengthString).setVerticalAlignment('middle').setHorizontalAlignment('center').mergeVertically();
  }
}