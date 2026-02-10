/** @OnlyCurrentDoc */

const shiyuDefenseFrontier5SheetName = "Shiyu Defense - Frontier 5";
const minShiyuDefenseFrontier5TeamStrengthRange = "G4";
const maxShiyuDefenseFrontier5OptionsRange = "G5";
const recalculateShiyuDefenseFrontier5Checkbox = "G2";
const shiyuDefenseFrontier5DistinctTeamsRange = "A2:D";
const shiyuDefenseFrontier5BuffsRangeAddress = "F8:G";

var shiyuDefenseFrontier5Team1BuffExpressions = [];
var shiyuDefenseFrontier5Team2BuffExpressions = [];
var shiyuDefenseFrontier5Team3BuffExpressions = [];

var _shiyuDefenseFrontier5Sheet;

function getShiyuDefenseFrontier5Sheet() {
  if (!_shiyuDefenseFrontier5Sheet) {
    _shiyuDefenseFrontier5Sheet = getSpreadsheet().getSheetByName(shiyuDefenseFrontier5SheetName);
  }
  return _shiyuDefenseFrontier5Sheet;
}

function initalizeShiyuDefenseFrontier5BuffExpressions() {
  shiyuDefenseFrontier5Team1BuffExpressions = [];
  shiyuDefenseFrontier5Team2BuffExpressions = [];
  shiyuDefenseFrontier5Team3BuffExpressions = [];
  const shiyuDefenseFrontier5BuffsRange = getShiyuDefenseFrontier5Sheet().getRange(shiyuDefenseFrontier5BuffsRangeAddress);
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
  const minShiyuDefenseFrontier5TeamStrength = getShiyuDefenseFrontier5Sheet().getRange(minShiyuDefenseFrontier5TeamStrengthRange).getValue();
  const allTeams = getAllTeams(minShiyuDefenseFrontier5TeamStrength);
  const teamTriples = computeBestDistinctTeamTriples(allTeams, shiyuDefenseFrontier5Team1BuffExpressions, shiyuDefenseFrontier5Team2BuffExpressions, shiyuDefenseFrontier5Team3BuffExpressions);
  const sortedTriples = teamTriples.sort((triple1, triple2) => triple2.minStrength() - triple1.minStrength() || triple2.totalStrength() - triple1.totalStrength());
  updateShiyuDefenseFrontier5DistinctTeamsSheet(sortedTriples);
}

function clearShiyuDefenseFrontier5Teams() {
  const shiyuDefenseFrontier5DistinctTeams = getShiyuDefenseFrontier5Sheet().getRange(shiyuDefenseFrontier5DistinctTeamsRange);
  shiyuDefenseFrontier5DistinctTeams.clearContent().breakApart();
}

function updateShiyuDefenseFrontier5DistinctTeamsSheet(teamTriples) {
  const sheet = getShiyuDefenseFrontier5Sheet();
  const shiyuDefenseFrontier5DistinctTeams = sheet.getRange(shiyuDefenseFrontier5DistinctTeamsRange);
  const shiyuDefenseFrontier5TeamsRow = shiyuDefenseFrontier5DistinctTeams.getRow();
  const shiyuDefenseFrontier5TeamsColumn = shiyuDefenseFrontier5DistinctTeams.getColumn();
  const maxShiyuDefenseFrontier5Options = sheet.getRange(maxShiyuDefenseFrontier5OptionsRange).getValue();

  if (teamTriples.length == 0) {
    sheet.getRange(shiyuDefenseFrontier5TeamsRow, shiyuDefenseFrontier5TeamsColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross()
  }

  var outputValues = [];
  var limit = Math.min(teamTriples.length, maxShiyuDefenseFrontier5Options);

  for (var i = 0; i < limit; i++) {
    var teamTriple = teamTriples[i];
    var team1 = teamTriple.team1;
    var team2 = teamTriple.team2;
    var team3 = teamTriple.team3;
    var strengthString =
      team1.strength +  " + " + teamTriple.team1Bonus + " \n+ " +
      team2.strength + " + " + teamTriple.team2Bonus + "\n= " +
      team3.strength + " + " + teamTriple.team3Bonus + "\n= " + teamTriple.totalStrength() + " (min=" + teamTriple.minStrength() + ")";

    outputValues.push([...team1.characters, strengthString]);
    outputValues.push([...team2.characters, ""]);
    outputValues.push([...team3.characters, ""]);
    outputValues.push(["", "", "", ""]);
  }

  if (outputValues.length > 0) {
    sheet.getRange(shiyuDefenseFrontier5TeamsRow, shiyuDefenseFrontier5TeamsColumn, outputValues.length, 4).setValues(outputValues);

    var strengthColIndex = shiyuDefenseFrontier5TeamsColumn + 3;
    sheet.getRange(shiyuDefenseFrontier5TeamsRow, strengthColIndex, outputValues.length, 1)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');

    // Optimization: Batch merge operations to avoid loop overhead
    sheet.getRange(shiyuDefenseFrontier5TeamsRow, strengthColIndex, 3, 1).mergeVertically();

    if (limit > 1) {
      var templateRange = sheet.getRange(shiyuDefenseFrontier5TeamsRow, strengthColIndex, 4, 1);
      var targetRange = sheet.getRange(shiyuDefenseFrontier5TeamsRow + 4, strengthColIndex, (limit - 1) * 4, 1);
      templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }
}
