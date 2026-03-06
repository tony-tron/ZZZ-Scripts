const assert = require('assert');

// Mock SpreadsheetApp
class Range {
  constructor(row, col, numRows, numCols) {
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
    this.values = null;
    this.alignments = null;
    this.valignments = null;
    this.merged = false;
  }
  setValue(val) {
    global.apiCalls++;
    this.values = [[val]];
    return this;
  }
  setValues(vals) {
    global.apiCalls++;
    this.values = vals;
    return this;
  }
  setHorizontalAlignment(val) {
    global.apiCalls++;
    this.alignments = val;
    return this;
  }
  setVerticalAlignment(val) {
    global.apiCalls++;
    this.valignments = val;
    return this;
  }
  mergeVertically() {
    global.apiCalls++;
    this.merged = true;
    return this;
  }
  mergeAcross() {
    global.apiCalls++;
    this.merged = true;
    return this;
  }
  copyTo(target, pasteType, flag) {
    global.apiCalls++;
    return this;
  }
}

class Sheet {
  getRange(row, col, numRows = 1, numCols = 1) {
    global.apiCalls++;
    return new Range(row, col, numRows, numCols);
  }
}

global.SpreadsheetApp = {
  CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT' }
};

const sheet = new Sheet();

// Mock data
const teamTriples = [];
for (let i = 0; i < 10; i++) {
  teamTriples.push({
    team1: { strength: 10, characters: ['A', 'B', 'C'] },
    team2: { strength: 20, characters: ['D', 'E', 'F'] },
    team3: { strength: 30, characters: ['G', 'H', 'I'] },
    team1Bonus: 1, team2Bonus: 2, team3Bonus: 3,
    team1ChosenBonus: 1, team2ChosenBonus: 2, team3ChosenBonus: 3,
    team1ChosenBonusName: 'x', team2ChosenBonusName: 'y', team3ChosenBonusName: 'z',
    totalStrength: () => 60, minStrength: () => 10
  });
}

// Old implementation
function oldImplementation() {
  global.apiCalls = 0;
  const startRow = 2, startColumn = 1, maxOptions = 10;

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
  return global.apiCalls;
}

// New implementation
function newImplementation() {
  global.apiCalls = 0;
  const startRow = 2, startColumn = 1, maxOptions = 10;

  var outputValues = [];
  var limit = Math.min(teamTriples.length, maxOptions);

  for (var i = 0; i < limit; i++) {
    var teamTriple = teamTriples[i];
    var team1 = teamTriple.team1;
    var team2 = teamTriple.team2;
    var team3 = teamTriple.team3;
    var strengthString =
      team1.strength + " + " + teamTriple.team1Bonus + " + " + teamTriple.team1ChosenBonus + "\n+ " +
      team2.strength + " + " + teamTriple.team2Bonus + " + " + teamTriple.team2ChosenBonus + "\n+ " +
      team3.strength + " + " + teamTriple.team3Bonus + " + " + teamTriple.team3ChosenBonus + "\n= " +
      teamTriple.totalStrength() + " (min= " + teamTriple.minStrength() + ")";

    outputValues.push([...team1.characters, strengthString, teamTriple.team1ChosenBonusName]);
    outputValues.push([...team2.characters, "", teamTriple.team2ChosenBonusName]);
    outputValues.push([...team3.characters, "", teamTriple.team3ChosenBonusName]);
    outputValues.push(["", "", "", "", ""]);
  }

  if (outputValues.length > 0) {
    sheet.getRange(startRow, startColumn, outputValues.length, 5).setValues(outputValues);

    var strengthColIndex = startColumn + 3;
    sheet.getRange(startRow, strengthColIndex, outputValues.length, 1)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');

    sheet.getRange(startRow, strengthColIndex, 3, 1).mergeVertically();

    var bonusColIndex = startColumn + 4;
    sheet.getRange(startRow, bonusColIndex, outputValues.length, 1)
      .setHorizontalAlignment('center');

    if (limit > 1) {
      var templateRange = sheet.getRange(startRow, strengthColIndex, 4, 1);
      var targetRange = sheet.getRange(startRow + 4, strengthColIndex, (limit - 1) * 4, 1);
      templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }

  return global.apiCalls;
}

const oldCalls = oldImplementation();
const newCalls = newImplementation();

console.log(`Old Implementation API Calls: ${oldCalls}`);
console.log(`New Implementation API Calls: ${newCalls}`);
console.log(`Improvement: ${Math.round((oldCalls - newCalls) / oldCalls * 100)}% reduction in API calls`);
