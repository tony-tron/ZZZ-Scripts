/** @OnlyCurrentDoc */

const deadlyAssaultSheetName = "Deadly Assault";
const recalculateDeadlyAssaultCheckbox = "H2";

function getDeadlyAssaultSheet() {
  return getSpreadsheet().getSheetByName(deadlyAssaultSheetName);
}

function getDeadlyAssaultDistinctTeamsRange() {
  return getDeadlyAssaultSheet().getRange("A2:E");
}

function getDeadlyAssaultBuffsRange() {
  return getDeadlyAssaultSheet().getRange("G8:H");
}

var deadlyAssaultTeam1BuffExpressions = [];
var deadlyAssaultTeam2BuffExpressions = [];
var deadlyAssaultTeam3BuffExpressions = [];
var deadlyAssaultBuffOptions = []; // Has `name` and `expression` fields.

function initalizeDeadlyAssaultBuffExpressions() {
  deadlyAssaultTeam1BuffExpressions = [];
  deadlyAssaultTeam2BuffExpressions = [];
  deadlyAssaultTeam3BuffExpressions = [];
  deadlyAssaultBuffOptions = [];

  const buffNamesAndExpressions = getDeadlyAssaultBuffsRange().getValues();

  var expression;
  var r = 0;
  var name;
  // Team 1 buffs first.
  for (; r < buffNamesAndExpressions.length; r++) {
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    deadlyAssaultTeam1BuffExpressions.push(expression);
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
    deadlyAssaultTeam2BuffExpressions.push(expression);
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
    deadlyAssaultTeam3BuffExpressions.push(expression);
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
  // Additional buff options.
  for (; r < buffNamesAndExpressions.length; r++) {
    name = buffNamesAndExpressions[r][0];
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    if (name == null || name == "") {
      name = "Buff " + (deadlyAssaultBuffOptions.length + 1);
    }
    deadlyAssaultBuffOptions.push({
      name : name,
      expression : expression,
    });
  }
}

function updateDeadlyAssaultSheet() {
  initalizeDeadlyAssaultBuffExpressions();

  const sheet = getDeadlyAssaultSheet();
  const minDeadlyAssaultTeamStrength = sheet.getRange("H4").getValue();
  const maxDeadlyAssaultOptions = sheet.getRange("H5").getValue();

  const deadlyAssaultDistinctTeams = getDeadlyAssaultDistinctTeamsRange();

  clearDeadlyAssaultTeams();
  const allTeams = getAllTeams(minDeadlyAssaultTeamStrength);
  const teamTriples = computeBestDistinctTeamTriples(allTeams, deadlyAssaultTeam1BuffExpressions, deadlyAssaultTeam2BuffExpressions, deadlyAssaultTeam3BuffExpressions, deadlyAssaultBuffOptions);
  const sortedTriples = teamTriples.sort((triple1, triple2) => triple2.minStrength() - triple1.minStrength() || triple2.totalStrength() - triple1.totalStrength());
  updateDeadlyAssaultDistinctTeamsSheet(sortedTriples, maxDeadlyAssaultOptions);
}

function clearDeadlyAssaultTeams() {
  getDeadlyAssaultDistinctTeamsRange().clearContent().breakApart();
}

function updateDeadlyAssaultDistinctTeamsSheet(teamTriples, maxDeadlyAssaultOptions) {
  const sheet = getDeadlyAssaultSheet();
  const distinctTeamsRange = getDeadlyAssaultDistinctTeamsRange();
  const deadlyAssaultTeamsRow = distinctTeamsRange.getRow();
  const deadlyAssaultTeamsColumn = distinctTeamsRange.getColumn();

  if (teamTriples.length == 0) {
    sheet.getRange(deadlyAssaultTeamsRow, deadlyAssaultTeamsColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross()
  }
  for (var i = 0; i < teamTriples.length && i < maxDeadlyAssaultOptions; i++) {
    var teamTriple = teamTriples[i];
    var team1 = teamTriple.team1;
    var team2 = teamTriple.team2;
    var team3 = teamTriple.team3;
    var strengthString =
      team1.strength + " + " + teamTriple.team1Bonus + " + " + teamTriple.team1ChosenBonus + "\n+ " +
      team2.strength + " + " + teamTriple.team2Bonus + " + " + teamTriple.team2ChosenBonus + "\n+ " +
      team3.strength + " + " + teamTriple.team3Bonus + " + " + teamTriple.team3ChosenBonus + "\n= " +
      teamTriple.totalStrength() + " (min= " + teamTriple.minStrength() + ")";
    sheet.getRange(deadlyAssaultTeamsRow + i * 4, deadlyAssaultTeamsColumn, 1, 3)
      .setValues([team1.characters]);
    sheet.getRange(deadlyAssaultTeamsRow + 1 + i * 4, deadlyAssaultTeamsColumn, 1, 3)
      .setValues([team2.characters]);
    sheet.getRange(deadlyAssaultTeamsRow + 2 + i * 4, deadlyAssaultTeamsColumn, 1, 3)
      .setValues([team3.characters]);
    sheet.getRange(deadlyAssaultTeamsRow + i * 4, deadlyAssaultTeamsColumn + 3, 3, 1)
      .setValue(strengthString).setVerticalAlignment('middle').setHorizontalAlignment('center').mergeVertically();
    sheet.getRange(deadlyAssaultTeamsRow + i * 4, deadlyAssaultTeamsColumn + 4, 3, 1)
      .setValues([[teamTriple.team1ChosenBonusName], [teamTriple.team2ChosenBonusName], [teamTriple.team3ChosenBonusName]])
      .setHorizontalAlignment('center');
  }
}
