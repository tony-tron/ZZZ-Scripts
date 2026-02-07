/** @OnlyCurrentDoc */

const distinctTeamsSheetName = "Distinct Teams";
const recalculateDistinctTeamsCheckbox = "H2";

// Global variables to store the parsed buff data
var buffExpressionsList = []; // An array of arrays, e.g., [ [slot1 buffs], [slot2 buffs] ]
var buffOptions = []; // The list of "chosen" buffs

/**
 * Reads the buff range and populates the global `buffExpressionsList` and `buffOptions`.
 */
function initalizeBuffExpressions(buffsRange) {
  // Reset global variables
  buffExpressionsList = [];
  buffOptions = [];
  
  const buffNamesAndExpressions = buffsRange.getValues();

  let currentBuffList = []; // Temp storage for the current slot's buffs
  let section = 'SLOT'; // We start by reading team/slot buffs
  let r = 0;

  // Loop through all rows in the buffsRange
  while (r < buffNamesAndExpressions.length) {
    const name = buffNamesAndExpressions[r][0];
    const expression = buffNamesAndExpressions[r][1];

    // Check for a header (any text in the name column, no text in the expression column)
    if (name && !expression) {
      // If this is a new slot header (e.g., "Team 2", "Slot 3")
      // and we've already collected buffs for the previous slot,
      // push the completed list.
      if (section === 'SLOT' && currentBuffList.length > 0) {
        buffExpressionsList.push(currentBuffList);
        currentBuffList = []; // Start a new list
      }

      // Check if this header is for the "options" section
      if (name.toLowerCase().includes('option') || name.toLowerCase().includes('additional')) {
        section = 'OPTIONS';
      } else {
        // Assume any other header is a new 'SLOT'
        section = 'SLOT';
      }
      
      // We've processed the header, move to the next row
      r++;
      continue;
    }

    // If not a header, check for data
    if (expression) {
      if (section === 'SLOT') {
        currentBuffList.push(expression);
      } else { // section === 'OPTIONS'
        buffOptions.push({
          name: name || `Buff ${buffOptions.length + 1}`,
          expression: expression,
        });
      }
    } else {
      // This is an empty row (no name, no expression)
      // This signals the end of a block
      if (section === 'SLOT' && currentBuffList.length > 0) {
        buffExpressionsList.push(currentBuffList);
        currentBuffList = [];
      }
    }
    
    r++;
  }

  // After the loop, push any remaining buffs from the last slot
  if (section === 'SLOT' && currentBuffList.length > 0) {
    buffExpressionsList.push(currentBuffList);
  }
}

/**
 * Main function to update the sheet.
 */
function updateDistinctTeamsSheet() {
  const distinctTeamsSheet = thisSpreadsheet.getSheetByName(distinctTeamsSheetName);
  const minTeamStrength = distinctTeamsSheet.getRange("H4").getValue();
  const maxOptions = distinctTeamsSheet.getRange("H5").getValue();

  const outputRangeHeader = distinctTeamsSheet.getRange("A1");
  const distinctTeamsOutputRange = distinctTeamsSheet.getRange("A2:E");
  const distinctTeamsOutputRow = distinctTeamsOutputRange.getRow();
  const distinctTeamsOutputCol = distinctTeamsOutputRange.getColumn();

  // Range for buff definitions
  const buffsRange = distinctTeamsSheet.getRange("G7:H");

  initalizeBuffExpressions(buffsRange); // Populates buffExpressionsList
  clearTeams(distinctTeamsOutputRange); // Clears the single output range
  
  const allTeams = getAllTeams(minTeamStrength);
  const k = buffExpressionsList.length; // Number of distinct teams to find

  // --- Input Validation ---
  if (k === 0) {
    distinctTeamsSheet.getRange(distinctTeamsOutputRow, distinctTeamsOutputCol).setValue("No buff expressions found. Check G7:H.");
    return;
  }
  if (k > 5) {
    distinctTeamsSheet.getRange(distinctTeamsOutputRow, distinctTeamsOutputCol).setValue(`Found ${k} buff sections. This script only supports up to 5.`);
    return;
  }

  // Call the new generic computation function
  const teams = computeBestDistinctTeams(allTeams, k);
  
  if (!teams) {
    distinctTeamsSheet.getRange(distinctTeamsOutputRow, distinctTeamsOutputCol).setValue(`Error: computeBestDistinctTeams for k=${k} not found.`);
    return;
  }
  
  // Sort results
  const sortedTeams = teams.sort((a, b) => b.minStrength() - a.minStrength() || b.totalStrength() - a.totalStrength());
  
  // Call the new generic update function
  updateSheetWithTeams(sortedTeams, k, distinctTeamsSheet, maxOptions, outputRangeHeader, distinctTeamsOutputRow, distinctTeamsOutputCol);
}

/**
 * NEW: Helper function to call the correct computation logic.
 * Relies on functions from team_optimizer.js
 */
function computeBestDistinctTeams(allTeams, k) {
  // buffExpressionsList and buffOptions are global and will be
  // used by the functions below.
  switch (k) {
    case 1:
      // Note: Assumes computeBestDistinctTeamSingles exists in team_optimizer.js
      return computeBestDistinctTeamSingles(allTeams, buffExpressionsList[0], buffOptions);
    case 2:
      return computeBestDistinctTeamPairs(allTeams, buffExpressionsList[0], buffExpressionsList[1], buffOptions);
    case 3:
      return computeBestDistinctTeamTriples(allTeams, buffExpressionsList[0], buffExpressionsList[1], buffExpressionsList[2], buffOptions);
    case 4:
      return computeBestDistinctTeamQuads(allTeams, buffExpressionsList, buffOptions);
    case 5:
      return computeBestDistinctTeamQuints(allTeams, buffExpressionsList, buffOptions);
    default:
      return [];
  }
}

function clearTeams(distinctTeamsOutputRange) {
  distinctTeamsOutputRange.clearContent().breakApart();
}

/**
 * Generic function to write teams to the sheet, formatted by k.
 */
function updateSheetWithTeams(teams, k, distinctTeamsSheet, maxOptions, outputRangeHeader, distinctTeamsOutputRow, distinctTeamsOutputCol) {
  if (teams.length === 0) {
    distinctTeamsSheet.getRange(distinctTeamsOutputRow, distinctTeamsOutputCol, 1, 5).setValue("No combination found, try lowering Min Strength").setHorizontalAlignment('center').mergeAcross();
    return;
  }

  outputRangeHeader.setValue(`Distinct Teams (${k})`);

  // k = number of teams per group
  const rowHeightPerGroup = k + 1; // k teams + 1 buffer row
  const teamCharacterCols = 3; // Columns for team characters
  const strengthColOffset = teamCharacterCols; // e.g., Col 4
  const chosenBuffColOffset = teamCharacterCols + 1; // e.g., Col 5

  for (let i = 0; i < teams.length && i < maxOptions; i++) {
    const teamGroup = teams[i];
    const currentRow = distinctTeamsOutputRow + (i * rowHeightPerGroup);

    let strengthString = "";
    let chosenBuffNames = [];
    let teamNames = [];

    // Build the data arrays for this group
    for (let j = 1; j <= k; j++) {
      const team = teamGroup[`team${j}`];
      const slotBonus = teamGroup[`team${j}Bonus`];
      
      // Handle 'chosen' buffs, which may not exist for k=1 or k=2
      const chosenBonus = teamGroup[`team${j}ChosenBonus`] || 0;
      const chosenName = teamGroup[`team${j}ChosenBonusName`] || "";
      
      teamNames.push(team.characters);
      chosenBuffNames.push([chosenName]);
      
      // Build strength string
      strengthString += `${team.strength} + ${slotBonus}`;
      if (chosenBonus > 0) {
        strengthString += ` + ${chosenBonus}`;
      }
      strengthString += "\n";
    }
    
    strengthString += `= ${teamGroup.totalStrength()} (min= ${teamGroup.minStrength()})`;

    // --- Write data to the sheet ---
    
    // 1. Write Team Characters
    // e.g., getRange(A2, 3 rows, 3 cols) for k=3
    distinctTeamsSheet.getRange(currentRow, distinctTeamsOutputCol, k, teamCharacterCols)
      .setValues(teamNames);
    
    // 2. Write Strength String
    // e.g., getRange(D2, 3 rows, 1 col) for k=3
    distinctTeamsSheet.getRange(currentRow, distinctTeamsOutputCol + strengthColOffset, k, 1)
      .setValue(strengthString)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .mergeVertically();

    // 3. Write Chosen Buff Names (if they exist on the object)
    if (teamGroup.hasOwnProperty('team1ChosenBonusName')) {
      distinctTeamsSheet.getRange(currentRow, distinctTeamsOutputCol + chosenBuffColOffset, k, 1)
        .setValues(chosenBuffNames)
        .setHorizontalAlignment('center');
    }
  }
}