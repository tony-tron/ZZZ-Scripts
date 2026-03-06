const { performance } = require('perf_hooks');

// Mock SpreadsheetApp
class MockRange {
  constructor(row, col, numRows, numCols) {
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
    this.values = [];
    this.merged = false;
  }
  setValues(values) {
    global.apiCalls++;
    this.values = values;
    return this;
  }
  setValue(value) {
    global.apiCalls++;
    this.values = [[value]];
    return this;
  }
  setVerticalAlignment() { global.apiCalls++; return this; }
  setHorizontalAlignment() { global.apiCalls++; return this; }
  mergeVertically() { global.apiCalls++; this.merged = true; return this; }
  mergeAcross() { global.apiCalls++; this.merged = true; return this; }
  copyTo(target, pasteType, transpose) { global.apiCalls++; return this; }
}

class MockSheet {
  getRange(row, col, numRows = 1, numCols = 1) {
    global.apiCalls++;
    return new MockRange(row, col, numRows, numCols);
  }
}

global.SpreadsheetApp = {
  CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT' }
};

const sheet = new MockSheet();

function generateMockData(n) {
  const triples = [];
  for (let i = 0; i < n; i++) {
    triples.push({
      team1: { strength: 10, characters: ['A', 'B', 'C'] },
      team2: { strength: 20, characters: ['D', 'E', 'F'] },
      team3: { strength: 30, characters: ['G', 'H', 'I'] },
      team1Bonus: 1, team1ChosenBonus: 2, team1ChosenBonusName: 'B1',
      team2Bonus: 1, team2ChosenBonus: 2, team2ChosenBonusName: 'B2',
      team3Bonus: 1, team3ChosenBonus: 2, team3ChosenBonusName: 'B3',
      totalStrength: () => 60,
      minStrength: () => 10
    });
  }
  return triples;
}

const teamTriples = generateMockData(20);
const startRow = 1;
const startColumn = 1;
const maxOptions = 20;

function originalImplementation() {
  for (var i = 0; i < teamTriples.length && i < maxOptions; i++) {
    var teamTriple = teamTriples[i];
    var team1 = teamTriple.team1;
    var team2 = teamTriple.team2;
    var team3 = teamTriple.team3;
    var strengthString = `${team1.strength} + ${teamTriple.team1Bonus} + ${teamTriple.team1ChosenBonus}
+ ${team2.strength} + ${teamTriple.team2Bonus} + ${teamTriple.team2ChosenBonus}
+ ${team3.strength} + ${teamTriple.team3Bonus} + ${teamTriple.team3ChosenBonus}
= ${teamTriple.totalStrength()} (min= ${teamTriple.minStrength()})`;
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

function optimizedImplementation() {
  const numOptions = Math.min(teamTriples.length, maxOptions);
  if (numOptions === 0) {
    sheet.getRange(startRow, startColumn, 1, 3).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross();
    return;
  }

  const data = [];
  for (let i = 0; i < numOptions; i++) {
    const teamTriple = teamTriples[i];
    const team1 = teamTriple.team1;
    const team2 = teamTriple.team2;
    const team3 = teamTriple.team3;
    const strengthString = `${team1.strength} + ${teamTriple.team1Bonus} + ${teamTriple.team1ChosenBonus}\n+ ${team2.strength} + ${teamTriple.team2Bonus} + ${teamTriple.team2ChosenBonus}\n+ ${team3.strength} + ${teamTriple.team3Bonus} + ${teamTriple.team3ChosenBonus}\n= ${teamTriple.totalStrength()} (min= ${teamTriple.minStrength()})`;

    data.push([team1.characters[0], team1.characters[1], team1.characters[2], strengthString, teamTriple.team1ChosenBonusName || '']);
    data.push([team2.characters[0], team2.characters[1], team2.characters[2], '', teamTriple.team2ChosenBonusName || '']);
    data.push([team3.characters[0], team3.characters[1], team3.characters[2], '', teamTriple.team3ChosenBonusName || '']);

    if (i < numOptions - 1) {
      data.push(['', '', '', '', '']);
    }
  }

  sheet.getRange(startRow, startColumn, data.length, 5).setValues(data);

  // Formatting
  sheet.getRange(startRow, startColumn + 3, data.length, 1).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.getRange(startRow, startColumn + 4, data.length, 1).setHorizontalAlignment('center');

  if (numOptions > 1) {
    sheet.getRange(startRow, startColumn + 3, 3, 1).mergeVertically();
    const templateRange = sheet.getRange(startRow, startColumn + 3, 4, 1);
    const targetRange = sheet.getRange(startRow + 4, startColumn + 3, (numOptions - 1) * 4, 1);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  } else {
    sheet.getRange(startRow, startColumn + 3, 3, 1).mergeVertically();
  }
}

global.apiCalls = 0;
const start1 = performance.now();
originalImplementation();
const end1 = performance.now();
const apiCalls1 = global.apiCalls;

global.apiCalls = 0;
const start2 = performance.now();
optimizedImplementation();
const end2 = performance.now();
const apiCalls2 = global.apiCalls;

console.log(`Original: ${apiCalls1} API calls, ${(end1 - start1).toFixed(2)} ms`);
console.log(`Optimized: ${apiCalls2} API calls, ${(end2 - start2).toFixed(2)} ms`);
