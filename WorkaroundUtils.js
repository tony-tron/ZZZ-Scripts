/** @OnlyCurrentDoc */

const readmeSheet = thisSpreadsheet.getSheetByName("README");
const latestVersionRange = readmeSheet.getRange("C1");
const extraSynergySheet = thisSpreadsheet.getSheetByName("Extra Synergy");
const extraSynergyRange = extraSynergySheet.getRange("A1");
const sumSynergySheet = thisSpreadsheet.getSheetByName("Sum 2: Synergy");
const synergyBonusRange = sumSynergySheet.getRange("J2");
const sumExtraSynergySheet = thisSpreadsheet.getSheetByName("Sum 3: Extra Synergy");
const extraSynergyBonusRange = sumExtraSynergySheet.getRange("K2");

function refreshLatestVersion() {
  refreshFormulasForRanges([latestVersionRange],
  ['=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1PdatbmxA9f1MXNmv4XCn9BUWFl8ZGVi776Px4VzuwV4/edit", "\'Version History\'!A1:B")']);
}

function refreshAllCustomFormulas(fastRefresh = false) {
  const rangesToRefresh = [
    extraSynergyRange,
    synergyBonusRange,
    extraSynergyBonusRange,
  ];

  const formulasInRefreshedRanges = [
    '=MAP(QUERY(Characters!A2:A, "SELECT A WHERE A IS NOT NULL"), QUERY(Characters!A2:L, "SELECT L WHERE A IS NOT NULL"), LAMBDA(char, aa_query, {char, IFERROR(TRANSPOSE(QUERY(Characters!A2:BG, "SELECT A WHERE "&QUERY_VARIABLE_NAMES_TO_COLUMNS(aa_query)&"AND NOT A=\'"&char&"\'", 0)))}))',
    '=CALCULATE_SYNERGY_BUFFS(K2:S)',
    '=CALCULATE_BUFFS({F2:F, G2:G, H2:H, ARRAYFORMULA(IF(ISBLANK(F2:F), , MAP(F2:F, G2:G, H2:H, LAMBDA(char1, char2, char3, "(" & IFNOTBLANK(VLOOKUP(char1, Characters!$A$2:$T, 14, FALSE), "0") & ") + (" & IFNOTBLANK(VLOOKUP(char2, Characters!$A$2:$BD, 14, FALSE), "0") & ") + (" & IFNOTBLANK(VLOOKUP(char3, Characters!$A$2:$T, 14, FALSE), "0") & ")"))))})',
  ]

  refreshFormulasForRanges(rangesToRefresh, formulasInRefreshedRanges, fastRefresh);
}

function refreshFormulasForRanges(rangesToRefresh, formulasInRefreshedRanges, fastRefresh = false) {
  // 1. Clear all ranges
  rangesToRefresh.forEach(range => {
    range.clearContent();
  });

  // 2. First Flush: Ensures all ranges are actually cleared in the sheet
  if (!fastRefresh) {
    SpreadsheetApp.flush();
  }

  // 3. Restore all formulas
  rangesToRefresh.forEach((range, index) => {
    range.setFormula(formulasInRefreshedRanges[index]);
  });

  // 4. Second Flush: Ensures all formulas are re-applied and begin recalculating
  SpreadsheetApp.flush();
}