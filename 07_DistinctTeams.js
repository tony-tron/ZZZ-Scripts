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
  const distinctTeamsSheet = getSpreadsheet().getSheetByName(distinctTeamsSheetName);
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
  const numGroups = Math.min(teams.length, maxOptions);
  const totalRows = numGroups * rowHeightPerGroup;

  if (totalRows === 0) return;

  const teamCharacterCols = 3; // Columns for team characters
  const strengthColIndex = teamCharacterCols; // Index 3
  const chosenBuffColIndex = teamCharacterCols + 1; // Index 4

  // Initialize the output array with empty strings
  // 5 columns: Char1, Char2, Char3, StrengthString, ChosenBuffName
  const outputValues = new Array(totalRows).fill(null).map(() => ["", "", "", "", ""]);

  for (let i = 0; i < numGroups; i++) {
    const teamGroup = teams[i];
    const groupStartRowIndex = i * rowHeightPerGroup; // Index in outputValues array
    const currentRow = distinctTeamsOutputRow + groupStartRowIndex;

    let strengthString = "";

    // Build the data arrays for this group
    for (let j = 1; j <= k; j++) {
      const team = teamGroup[`team${j}`];
      const slotBonus = teamGroup[`team${j}Bonus`];
      
      // Handle 'chosen' buffs, which may not exist for k=1 or k=2
      const chosenBonus = teamGroup[`team${j}ChosenBonus`] || 0;
      const chosenName = teamGroup[`team${j}ChosenBonusName`] || "";
      
      // Fill characters (Columns 0, 1, 2)
      outputValues[groupStartRowIndex + j - 1][0] = team.characters[0];
      outputValues[groupStartRowIndex + j - 1][1] = team.characters[1];
      outputValues[groupStartRowIndex + j - 1][2] = team.characters[2];

      // Fill chosen buff name (Column 4)
      outputValues[groupStartRowIndex + j - 1][chosenBuffColIndex] = chosenName;
      
      // Build strength string
      strengthString += `${team.strength} + ${slotBonus}`;
      if (chosenBonus > 0) {
        strengthString += ` + ${chosenBonus}`;
      }
      strengthString += "\n";
    }
    
    strengthString += `= ${teamGroup.totalStrength()} (min= ${teamGroup.minStrength()})`;

    // Assign strength string to the first row of the group, column index 3 (4th col)
    outputValues[groupStartRowIndex][strengthColIndex] = strengthString;

    // Apply formatting (Merging and Alignment)
    // Strength Column (Column 4 -> Index 3 + OutputCol)
    distinctTeamsSheet.getRange(currentRow, distinctTeamsOutputCol + strengthColIndex, k, 1)
      .mergeVertically()
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');

    // Chosen Buff Column (Column 5 -> Index 4 + OutputCol)
    distinctTeamsSheet.getRange(currentRow, distinctTeamsOutputCol + chosenBuffColIndex, k, 1)
      .setHorizontalAlignment('center');
  }

  // Write all data at once
  distinctTeamsSheet.getRange(distinctTeamsOutputRow, distinctTeamsOutputCol, totalRows, 5)
    .setValues(outputValues);
}