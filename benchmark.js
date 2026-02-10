
// Mocking Google Apps Script Environment
const SpreadsheetApp = {
  CopyPasteType: {
    PASTE_FORMAT: 'PASTE_FORMAT'
  },
  callCount: 0,
  resetCount: () => { SpreadsheetApp.callCount = 0; }
};

class Range {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  setValues(values) {
    SpreadsheetApp.callCount++;
    // In a real mock, we would store values, but here we just count calls.
    return this;
  }

  setValue(value) {
    SpreadsheetApp.callCount++;
    return this;
  }

  setVerticalAlignment(align) {
    SpreadsheetApp.callCount++;
    return this;
  }

  setHorizontalAlignment(align) {
    SpreadsheetApp.callCount++;
    return this;
  }

  mergeVertically() {
    SpreadsheetApp.callCount++;
    return this;
  }

  copyTo(destination, type, transposed) {
    SpreadsheetApp.callCount++;
    // Simulate copying format (merge state)
    return this;
  }
}

class Sheet {
  constructor(name) {
    this.name = name;
  }

  getRange(row, col, numRows, numCols) {
    SpreadsheetApp.callCount++; // getRange is an API call
    if (numRows === undefined) numRows = 1;
    if (numCols === undefined) numCols = 1;
    return new Range(this, row, col, numRows, numCols);
  }
}

// Setup
const sheet = new Sheet("TestSheet");
const startRow = 2;
const startColumn = 1;
const count = 100; // Simulate 100 teams
const outputValues = new Array(count * 3).fill(["val1", "val2", "val3", "val4"]);


// ---------------------------------------------------------
// Original Method (Simulating the loop)
// ---------------------------------------------------------
function originalMethod() {
  console.log("Running Original Method...");
  SpreadsheetApp.resetCount();

  if (count > 0) {
    sheet.getRange(startRow, startColumn, count * 3, 4).setValues(outputValues);
    sheet.getRange(startRow, startColumn + 3, count * 3, 1).setVerticalAlignment('middle').setHorizontalAlignment('center');

    for (var i = 0; i < count; i++) {
      sheet.getRange(startRow + i * 3, startColumn + 3, 2, 1).mergeVertically();
    }
  }

  console.log(`Original Method Calls: ${SpreadsheetApp.callCount}`);
  return SpreadsheetApp.callCount;
}


// ---------------------------------------------------------
// Optimized Method (Using copyTo)
// ---------------------------------------------------------
function optimizedMethod() {
  console.log("Running Optimized Method...");
  SpreadsheetApp.resetCount();

  if (count > 0) {
    // 1. Write values (1 call to getRange, 1 call to setValues)
    sheet.getRange(startRow, startColumn, count * 3, 4).setValues(outputValues);

    // 2. Set alignment (1 call to getRange, 2 calls for alignment)
    sheet.getRange(startRow, startColumn + 3, count * 3, 1).setVerticalAlignment('middle').setHorizontalAlignment('center');

    // 3. Merge optimization
    // Merge the first block manually
    // (1 call to getRange, 1 call to mergeVertically)
    sheet.getRange(startRow, startColumn + 3, 2, 1).mergeVertically();

    // If more blocks, copy format
    if (count > 1) {
      // (2 calls to getRange, 1 call to copyTo)
      const templateRange = sheet.getRange(startRow, startColumn + 3, 3, 1);
      const targetRange = sheet.getRange(startRow + 3, startColumn + 3, (count - 1) * 3, 1);
      templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }

  console.log(`Optimized Method Calls: ${SpreadsheetApp.callCount}`);
  return SpreadsheetApp.callCount;
}

// Run Benchmark
const originalCalls = originalMethod();
const optimizedCalls = optimizedMethod();

console.log("\n--- Results ---");
console.log(`Original Calls (N=${count}): ${originalCalls}`);
console.log(`Optimized Calls (N=${count}): ${optimizedCalls}`);

if (optimizedCalls < originalCalls) {
  console.log("SUCCESS: Optimization reduced API calls!");
} else {
  console.log("FAILURE: Optimization did not reduce API calls.");
}
