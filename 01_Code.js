/** @OnlyCurrentDoc */

// Removed global spreadsheet read. Use getSpreadsheet() instead.

var _spreadsheet;

function getSpreadsheet() {
  if (!_spreadsheet) {
    _spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _spreadsheet;
}

function init() {
  initializeAllTeamsAndBuffParams();
}

function onOpen() {
  refreshLatestVersion();
  init();

  const menuEntries = [
      {name: 'Recompute Shiyu Defense Sheets', functionName: 'updateShiyuDefenseSheets'},
      {name: 'Recompute Deadly Assault Sheet', functionName: 'updateDeadlyAssaultSheet'},
      {name: 'Recompute Distinct Teams Sheet', functionName: 'updateDistinctTeamsSheet'},
      {name: 'Recompute Who To Build Sheet', functionName: "updateWhoToBuildSheet"},
      {name: 'Recompute Tier List Sheet', functionName: "updateTierListSheet"},
      {name: 'Refresh Formulas (if you see errors)', functionName: 'refreshAllCustomFormulas'},
    ];
  getSpreadsheet().addMenu('Advanced Actions', menuEntries);
}

function onEdit(e) {
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  if (sheetName === sortedTeamsSheetName && e.range.getA1Notation() === refreshFormulasCheckbox) {
    getSpreadsheet().toast('Refreshing...');
    refreshAllCustomFormulas(true);
  }
  if (sheetName === shiyuDefenseFrontier4SheetName && e.range.getA1Notation() === recalculateShiyuDefenseFrontier4Checkbox) {
    getSpreadsheet().toast('Recalculating...');
    updateShiyuDefenseFrontier4Sheet();
  }
  if (sheetName === shiyuDefenseFrontier5SheetName && e.range.getA1Notation() === recalculateShiyuDefenseFrontier5Checkbox) {
    getSpreadsheet().toast('Recalculating...');
    updateShiyuDefenseFrontier5Sheet();
  }
  if (sheetName === deadlyAssaultSheetName && e.range.getA1Notation() === recalculateDeadlyAssaultCheckbox) {
    getSpreadsheet().toast('Recalculating...');
    updateDeadlyAssaultSheet();
  }
  if (sheetName === distinctTeamsSheetName && e.range.getA1Notation() === recalculateDistinctTeamsCheckbox) {
    getSpreadsheet().toast('Recalculating...');
    updateDistinctTeamsSheet();
  }
  if (sheetName === tierListSheetName && e.range.getA1Notation() === recalculateTierListCheckbox) {
    getSpreadsheet().toast('Recalculating...');
    updateTierListSheet();
  }
}

function updateShiyuDefenseSheets() {
  updateShiyuDefenseFrontier4Sheet();
  updateShiyuDefenseFrontier5Sheet();
}
