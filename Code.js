/** @OnlyCurrentDoc */

const thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function init() {
  charactersColumns = initCharactersColumns();
  charsToBuffParams = initCharsToBuffParams();
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
  thisSpreadsheet.addMenu('Advanced Actions', menuEntries);
}

function onEdit(e) {
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  if (sheetName === sortedTeamsSheet.getName() && e.range.getA1Notation() === refreshFormulasCheckbox) {
    SpreadsheetApp.getActive().toast('Refreshing...');
    refreshAllCustomFormulas(true);
  }
  if (sheetName === shiyuDefenseFrontier4Sheet.getName() && e.range.getA1Notation() === recalculateShiyuDefenseFrontier4Checkbox) {
    SpreadsheetApp.getActive().toast('Recalculating...');
    updateShiyuDefenseFrontier4Sheet();
  }
  if (sheetName === shiyuDefenseFrontier5Sheet.getName() && e.range.getA1Notation() === recalculateShiyuDefenseFrontier5Checkbox) {
    SpreadsheetApp.getActive().toast('Recalculating...');
    updateShiyuDefenseFrontier5Sheet();
  }
  if (sheetName === deadlyAssaultSheet.getName() && e.range.getA1Notation() === recalculateDeadlyAssaultCheckbox) {
    SpreadsheetApp.getActive().toast('Recalculating...');
    updateDeadlyAssaultSheet();
  }
  if (sheetName === distinctTeamsSheet.getName() && e.range.getA1Notation() === recalculateDistinctTeamsCheckbox) {
    SpreadsheetApp.getActive().toast('Recalculating...');
    updateDistinctTeamsSheet();
  }
  if (sheetName === tierListSheet.getName() && e.range.getA1Notation() === recalculateTierListCheckbox) {
    SpreadsheetApp.getActive().toast('Recalculating...');
    updateTierListSheet();
  }
}

function updateShiyuDefenseSheets() {
  updateShiyuDefenseFrontier4Sheet();
  updateShiyuDefenseFrontier5Sheet();
}