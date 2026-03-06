const fs = require('fs');
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
  constructor() {
    this.ranges = [];
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    global.apiCalls++;
    const range = new Range(row, col, numRows, numCols);
    this.ranges.push(range);
    return range;
  }
}

global.SpreadsheetApp = {
  CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT' }
};

const sheet = new Sheet();

// Mock data
const teamTriples = [];
for (let i = 0; i < 2; i++) {
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

// Load function
const code = fs.readFileSync('06_DeadlyAssault.js', 'utf8');
const updateDeadlyAssaultDistinctTeamsSheet = new Function('teamTriples', 'sheet', 'startRow', 'startColumn', 'maxOptions',
  code.match(/function updateDeadlyAssaultDistinctTeamsSheet[\s\S]*?^}/m)[0] + '\nreturn updateDeadlyAssaultDistinctTeamsSheet(teamTriples, sheet, startRow, startColumn, maxOptions);'
);

global.apiCalls = 0;
updateDeadlyAssaultDistinctTeamsSheet(teamTriples, sheet, 2, 1, 10);

console.log("Mock calls completed without errors.");
console.log(`Total API Calls: ${global.apiCalls}`);

const mainSetValuesRange = sheet.ranges.find(r => r.values && r.values.length > 2);
assert(mainSetValuesRange, "Should have a range with setValues call for multiple rows");
assert.equal(mainSetValuesRange.values.length, 8, "Should have 8 rows for 2 team triples");
assert.equal(mainSetValuesRange.values[0].length, 5, "Should have 5 columns");
assert.deepEqual(mainSetValuesRange.values[0].slice(0, 3), ['A', 'B', 'C'], "Row 0 should have Team 1 characters");
assert.equal(mainSetValuesRange.values[0][4], 'x', "Row 0 should have Team 1 bonus name");
assert.deepEqual(mainSetValuesRange.values[1].slice(0, 3), ['D', 'E', 'F'], "Row 1 should have Team 2 characters");
assert.equal(mainSetValuesRange.values[1][4], 'y', "Row 1 should have Team 2 bonus name");
assert.deepEqual(mainSetValuesRange.values[2].slice(0, 3), ['G', 'H', 'I'], "Row 2 should have Team 3 characters");
assert.equal(mainSetValuesRange.values[2][4], 'z', "Row 2 should have Team 3 bonus name");
assert.deepEqual(mainSetValuesRange.values[3], ['', '', '', '', ''], "Row 3 should be empty padding");

console.log("All assertions passed!");
