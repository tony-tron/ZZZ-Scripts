/** @OnlyCurrentDoc */

const shiyuDefenseFrontier4SheetName = "Shiyu Defense - Frontier 4";
const shiyuDefenseFrontier4DistinctTeamsA1 = "A2:D";
const minShiyuDefenseFrontier4TeamStrengthA1 = "G4";
const maxShiyuDefenseFrontier4OptionsA1 = "G5";
const recalculateShiyuDefenseFrontier4Checkbox = "G2";
const shiyuDefenseFrontier4BuffsRangeA1 = "F8:G";


function getShiyuDefenseFrontier4Context() {
  const sheet = getSpreadsheet().getSheetByName(shiyuDefenseFrontier4SheetName);
  const distinctTeamsRange = sheet.getRange(shiyuDefenseFrontier4DistinctTeamsA1);
  return {
    sheet: sheet,
    minStrength: sheet.getRange(minShiyuDefenseFrontier4TeamStrengthA1).getValue(),
    maxOptions: sheet.getRange(maxShiyuDefenseFrontier4OptionsA1).getValue(),
    buffRange: sheet.getRange(shiyuDefenseFrontier4BuffsRangeA1),
    distinctTeamsRange: distinctTeamsRange,
    startRow: distinctTeamsRange.getRow(),
    startColumn: distinctTeamsRange.getColumn()
  };
}

function getShiyuDefenseFrontier4BuffExpressions(buffRange) {
  const buffNamesAndExpressions = buffRange.getValues();
  const team1Buffs = [];
  const team2Buffs = [];

  var expression;
  var r = 0;
  // Team 1 buffs first.
  for (; r < buffNamesAndExpressions.length; r++) {
    expression = buffNamesAndExpressions[r][1];
    if (expression == null || expression == "") {
      break;
    }
    team1Buffs.push(expression);
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
    team2Buffs.push(expression);
  }
  return { team1Buffs, team2Buffs };
}

function updateShiyuDefenseFrontier4Sheet() {
  const ctx = getShiyuDefenseFrontier4Context();
  const { team1Buffs, team2Buffs } = getShiyuDefenseFrontier4BuffExpressions(ctx.buffRange);

  clearShiyuDefenseFrontier4Teams(ctx.distinctTeamsRange);
  const allTeams = getAllTeams(ctx.minStrength);
  const teamPairs = computeBestDistinctTeamPairs(allTeams, team1Buffs, team2Buffs);
  const sortedPairs = teamPairs.sort((pair1, pair2) => pair2.minStrength() - pair1.minStrength() || pair2.totalStrength() - pair1.totalStrength());
  updateShiyuDefenseFrontier4DistinctTeamsSheet(sortedPairs, ctx.sheet, ctx.startRow, ctx.startColumn, ctx.maxOptions);
}

function clearShiyuDefenseFrontier4Teams(range) {
  range.clearContent().breakApart();
}

function updateShiyuDefenseFrontier4DistinctTeamsSheet(teamPairs, sheet, startRow, startColumn, maxOptions) {
  if (teamPairs.length == 0) {
    sheet.getRange(startRow, startColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross()
  }
  var outputValues = [];
  var count = 0;
  for (var i = 0; i < teamPairs.length && i < maxOptions; i++) {
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
    sheet.getRange(startRow, startColumn, count * 3, 4).setValues(outputValues);
    sheet.getRange(startRow, startColumn + 3, count * 3, 1).setVerticalAlignment('middle').setHorizontalAlignment('center');

    for (var i = 0; i < count; i++) {
      sheet.getRange(startRow + i * 3, startColumn + 3, 2, 1).mergeVertically();
    }
  }
}
