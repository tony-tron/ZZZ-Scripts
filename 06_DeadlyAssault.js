/** @OnlyCurrentDoc */

const deadlyAssaultSheetName = "Deadly Assault";
const recalculateDeadlyAssaultCheckbox = "H2";
const deadlyAssaultDistinctTeamsA1 = "A2:E";
const deadlyAssaultBuffsRangeA1 = "G8:H";
const minDeadlyAssaultTeamStrengthA1 = "H4";
const maxDeadlyAssaultOptionsA1 = "H5";

function getDeadlyAssaultContext() {
  const sheet = getSpreadsheet().getSheetByName(deadlyAssaultSheetName);
  const distinctTeamsRange = sheet.getRange(deadlyAssaultDistinctTeamsA1);
  return {
    sheet: sheet,
    minStrength: sheet.getRange(minDeadlyAssaultTeamStrengthA1).getValue(),
    maxOptions: sheet.getRange(maxDeadlyAssaultOptionsA1).getValue(),
    buffRange: sheet.getRange(deadlyAssaultBuffsRangeA1),
    distinctTeamsRange: distinctTeamsRange,
    startRow: distinctTeamsRange.getRow(),
    startColumn: distinctTeamsRange.getColumn()
  };
}

function getDeadlyAssaultBuffExpressions(buffRange) {
  const team1BuffExpressions = [];
  const team2BuffExpressions = [];
  const team3BuffExpressions = [];
  const buffOptions = [];

  const buffNamesAndExpressions = buffRange.getValues();

  let r = 0;
  // Team 1 buffs first.
  for (; r < buffNamesAndExpressions.length; r++) {
    const expression = buffNamesAndExpressions[r][1];
    if (!expression) {
      break;
    }
    team1BuffExpressions.push(expression);
  }
  // Skip over empty cells.
  for (; r < buffNamesAndExpressions.length; r++) {
    const name = buffNamesAndExpressions[r][0];
    const expression = buffNamesAndExpressions[r][1];
    if (!name && !expression) {
      continue;
    } else if (name) {
      // We reached the header for the next section.
      r++;
      break;
    }
  }
  // Team 2 buffs.
  for (; r < buffNamesAndExpressions.length; r++) {
    const expression = buffNamesAndExpressions[r][1];
    if (!expression) {
      break;
    }
    team2BuffExpressions.push(expression);
  }
  // Skip over empty cells.
  for (; r < buffNamesAndExpressions.length; r++) {
    const name = buffNamesAndExpressions[r][0];
    const expression = buffNamesAndExpressions[r][1];
    if (!name && !expression) {
      continue;
    } else if (name) {
      // We reached the header for the next section.
      r++;
      break;
    }
  }
  // Team 3 buffs.
  for (; r < buffNamesAndExpressions.length; r++) {
    const expression = buffNamesAndExpressions[r][1];
    if (!expression) {
      break;
    }
    team3BuffExpressions.push(expression);
  }
  // Skip over empty cells.
  for (; r < buffNamesAndExpressions.length; r++) {
    const name = buffNamesAndExpressions[r][0];
    const expression = buffNamesAndExpressions[r][1];
    if (!name && !expression) {
      continue;
    } else if (name) {
      // We reached the header for the next section.
      r++;
      break;
    }
  }
  // Additional buff options.
  for (; r < buffNamesAndExpressions.length; r++) {
    let name = buffNamesAndExpressions[r][0];
    const expression = buffNamesAndExpressions[r][1];
    if (!expression) {
      break;
    }
    if (!name) {
      name = "Buff " + (buffOptions.length + 1);
    }
    buffOptions.push({
      name : name,
      expression : expression,
    });
  }
  return { team1BuffExpressions, team2BuffExpressions, team3BuffExpressions, buffOptions };
}

function updateDeadlyAssaultSheet() {
  const ctx = getDeadlyAssaultContext();
  const { team1BuffExpressions, team2BuffExpressions, team3BuffExpressions, buffOptions } = getDeadlyAssaultBuffExpressions(ctx.buffRange);

  clearDeadlyAssaultTeams(ctx.distinctTeamsRange);
  const allTeams = getAllTeams(ctx.minStrength);
  const teamTriples = computeBestDistinctTeamTriples(allTeams, team1BuffExpressions, team2BuffExpressions, team3BuffExpressions, buffOptions);
  const sortedTriples = teamTriples.sort((triple1, triple2) => triple2.minStrength() - triple1.minStrength() || triple2.totalStrength() - triple1.totalStrength());
  updateDeadlyAssaultDistinctTeamsSheet(sortedTriples, ctx.sheet, ctx.startRow, ctx.startColumn, ctx.maxOptions);
}

function clearDeadlyAssaultTeams(range) {
  const rangeToClear = range || getDeadlyAssaultContext().distinctTeamsRange;
  rangeToClear.clearContent().breakApart();
}

function updateDeadlyAssaultDistinctTeamsSheet(teamTriples, sheet, startRow, startColumn, maxOptions) {
  if (teamTriples.length == 0) {
    sheet.getRange(startRow, startColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross()
  }
  for (var i = 0; i < teamTriples.length && i < maxOptions; i++) {
    var teamTriple = teamTriples[i];
    var team1 = teamTriple.team1;
    var team2 = teamTriple.team2;
    var team3 = teamTriple.team3;
    var strengthString =
      team1.strength + " + " + teamTriple.team1Bonus + " + " + teamTriple.team1ChosenBonus + "\n+ " +
      team2.strength + " + " + teamTriple.team2Bonus + " + " + teamTriple.team2ChosenBonus + "\n+ " +
      team3.strength + " + " + teamTriple.team3Bonus + " + " + teamTriple.team3ChosenBonus + "\n= " +
      teamTriple.totalStrength() + " (min= " + teamTriple.minStrength() + ")";
    sheet.getRange(startRow + i * 4, startColumn, 1, 3)
      .setValues([team1.characters]);
    sheet.getRange(startRow + 1 + i * 4, startColumn, 1, 3)
      .setValues([team2.characters]);
    sheet.getRange(startRow + 2 + i * 4, startColumn, 1, 3)
      .setValues([team3.characters]);
    sheet.getRange(startRow + i * 4, startColumn + 3, 3, 1)
      .setValue(strengthString).setVerticalAlignment('middle').setHorizontalAlignment('center').mergeVertically();
    sheet.getRange(startRow + i * 4, startColumn + 4, 3, 1)
      .setValues([[teamTriple.team1ChosenBonusName], [teamTriple.team2ChosenBonusName], [teamTriple.team3ChosenBonusName]])
      .setHorizontalAlignment('center');
  }
}
