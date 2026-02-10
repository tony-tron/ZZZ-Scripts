/** @OnlyCurrentDoc */

// Removed global constants. Fetched dynamically.

var _readmeSheet;
var _latestVersionRange;
var _extraSynergySheet;
var _extraSynergyRange;
var _sumSynergySheet;
var _synergyBonusRange;
var _sumExtraSynergySheet;
var _extraSynergyBonusRange;

function getReadmeSheet() {
  if (!_readmeSheet) {
    _readmeSheet = getSpreadsheet().getSheetByName("README");
  }
  return _readmeSheet;
}

function getLatestVersionRange() {
  if (!_latestVersionRange) {
    _latestVersionRange = getReadmeSheet().getRange("C1");
  }
  return _latestVersionRange;
}

function getExtraSynergySheet() {
  if (!_extraSynergySheet) {
    _extraSynergySheet = getSpreadsheet().getSheetByName("Extra Synergy");
  }
  return _extraSynergySheet;
}

function getExtraSynergyRange() {
  if (!_extraSynergyRange) {
    _extraSynergyRange = getExtraSynergySheet().getRange("A1");
  }
  return _extraSynergyRange;
}

function getSumSynergySheet() {
  if (!_sumSynergySheet) {
    _sumSynergySheet = getSpreadsheet().getSheetByName("Sum 2: Synergy");
  }
  return _sumSynergySheet;
}

function getSynergyBonusRange() {
  if (!_synergyBonusRange) {
    _synergyBonusRange = getSumSynergySheet().getRange("J2");
  }
  return _synergyBonusRange;
}

function getSumExtraSynergySheet() {
  if (!_sumExtraSynergySheet) {
    _sumExtraSynergySheet = getSpreadsheet().getSheetByName("Sum 3: Extra Synergy");
  }
  return _sumExtraSynergySheet;
}

function getExtraSynergyBonusRange() {
  if (!_extraSynergyBonusRange) {
    _extraSynergyBonusRange = getSumExtraSynergySheet().getRange("K2");
  }
  return _extraSynergyBonusRange;
}

function refreshLatestVersion() {
  refreshFormulasForRanges([getLatestVersionRange()],
  ['=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1PdatbmxA9f1MXNmv4XCn9BUWFl8ZGVi776Px4VzuwV4/edit", "\'Version History\'!A1:B")']);
}

function refreshAllCustomFormulas(fastRefresh = false) {
  _charsToBuffParams = null;
  teamCharsToTeamObjs = null;
  init();

  const rangesToRefresh = fastRefresh
  ? [getExtraSynergyBonusRange()]
  : [
    getExtraSynergyRange(),
    getSynergyBonusRange(),
    getExtraSynergyBonusRange(),
  ];

  const formulasInRefreshedRanges = [
    '=MAP(QUERY(Characters!A2:A, "SELECT A WHERE A IS NOT NULL"), QUERY(Characters!A2:L, "SELECT L WHERE A IS NOT NULL"), LAMBDA(char, aa_query, {char, IFERROR(TRANSPOSE(QUERY(Characters!A2:BG, "SELECT A WHERE "&QUERY_VARIABLE_NAMES_TO_COLUMNS(aa_query)&"AND NOT A=\'"&char&"\'", 0)))}))',
    '=CALCULATE_SYNERGY_BUFFS(K2:S)',
    '=CALCULATE_TEAM_BUFFS(F2:H)',
  ]

  refreshFormulasForRanges(rangesToRefresh, formulasInRefreshedRanges);
}

function refreshFormulasForRanges(rangesToRefresh, formulasInRefreshedRanges) {
  // 1. Clear all ranges
  rangesToRefresh.forEach(range => {
    range.clearContent();
  });

  // 2. First Flush: Ensures all ranges are actually cleared in the sheet
  SpreadsheetApp.flush();

  // 3. Restore all formulas
  rangesToRefresh.forEach((range, index) => {
    range.setFormula(formulasInRefreshedRanges[index]);
  });

  // 4. Second Flush: Ensures all formulas are re-applied and begin recalculating
  SpreadsheetApp.flush();
}
