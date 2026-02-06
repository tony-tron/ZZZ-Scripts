/** @OnlyCurrentDoc */

const whoToBuildSheet = thisSpreadsheet.getSheetByName("Who To Build");
const unbuiltCharactersRange = whoToBuildSheet.getRange("A2:B");
const unbuiltCharacters = unbuiltCharactersRange.getValues();
const outputRange = whoToBuildSheet.getRange("D2:G");
const outputClearFormatRange = whoToBuildSheet.getRange("D2:F");

/** Returns true if there is an error in the input, meaning we should abort the rest of the script. */
function checkAndOutputErrors(numCharactersConsidered) {
  const outputMessageRange = whoToBuildSheet.getRange(outputRange.getRow(), outputRange.getColumn());
  var outputMessage = "";
  var isError = false;
  if (numCharactersConsidered == 0) {
    outputMessage = "You must select at least one character from the left.";
    isError = true;
  } else if (numCharactersConsidered > 20) {
    outputMessage = "Please select 20 or fewer characters or it might time out.";
    isError = true;
  } else {
    const estimatedTime = secondsToString(roundToNearest15(13 + numCharactersConsidered * 17.0));
    outputMessage =
      "Please do not modify the spreadsheet or cancel the script while calculations\n" +
      "are being performed. This will take approximately " + estimatedTime + ".";
    isError = false;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
      isError ? "Error" : "Confirm",
      outputMessage,
      ui.ButtonSet.OK_CANCEL,
  );

  // Process the user's response.
  if (response === ui.Button.OK) {
    outputClearFormatRange.clearFormat().breakApart();
    outputRange.clearContent();
    outputMessageRange.setFontColor("#980000").setValue(outputMessage);
    return isError;
  } else {
    // User clicked Cancel or X, abort.
    isError = true;
  }

  return isError;
}

function displayOutputs(outputs) {
  outputs.sort((output1, output2) => output2.tierScore - output1.tierScore);

  const outputValues = [];
  const textStyles = [];
  const bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  const underline = SpreadsheetApp.newTextStyle().setUnderline(true).setFontSize(12).build();
  const mergeRowIndices = [];
  for (const output of outputs) {
    outputValues.push([
      output.character,
      output.numMetaTeams + " Meta Teams",
      "Strongest: " + output.maxTeamStrength,
      output.tier + " for you"
    ]);
    textStyles.push([bold, null, null]);
    mergeRowIndices.push(outputValues.length);
    outputValues.push(["Strongest Teams", null, null, null]);
    textStyles.push([underline, null, null]);
    for (const strongestTeam of output.strongestTeams) {
      outputValues.push(strongestTeam);
      textStyles.push([null, null, null]);
    }
    outputValues.push([null, null, null, null]);
    textStyles.push([null, null, null]);
  }
  for (var r = outputValues.length; r < outputRange.getNumRows(); r++) {
    outputValues.push([null, null, null, null]);
    textStyles.push([null, null, null]);
  }
  outputClearFormatRange.clearFormat().setTextStyles(textStyles);
  outputRange.clearContent().setValues(outputValues);
  for (const mergeRowIndex of mergeRowIndices) {
    whoToBuildSheet.getRange(
      outputRange.getRow() + mergeRowIndex,
      outputRange.getColumn(),
      1, outputRange.getNumColumns() - 1)
      .setHorizontalAlignment('center').mergeAcross();
  }
}

function updateWhoToBuildSheet() {
  const charactersToConsider = getCharactersToConsider();
  if (checkAndOutputErrors(charactersToConsider.length)) {
    return;
  }

  // Override the unbuilt characters formula while calculations are being performed,
  // we will set it back at the end.
  const unbuiltCharactersFormula = unbuiltCharactersRange.getFormula();
  unbuiltCharactersRange.setValues(unbuiltCharacters);
  unbuiltCharacters[0][0] = unbuiltCharactersFormula;
  for (var r = 1; r< unbuiltCharacters.length; r++) {
    unbuiltCharacters[r][0] = null;
  }

  const tempOnlyIncludeBuilt = true;
  const tempOnlyIncludeReleased = false;
  const oldSortedTeamsCheckboxValues = setSortedTeamsCheckboxesAndGetOldValuesToRestoreLater(
    tempOnlyIncludeBuilt, tempOnlyIncludeReleased);

  const tierListParams = getTierListParams();
  const outputs = [];
  setCharactersBuilt([charactersToConsider[0]], [true]);
  addOutputForCharacter(outputs, charactersToConsider[0], tierListParams);
  for (var i = 1; i < charactersToConsider.length; i++) {
    setCharactersBuilt([charactersToConsider[i-1], charactersToConsider[i]], [false, true]);
    addOutputForCharacter(outputs, charactersToConsider[i], tierListParams);
  }
  setCharactersBuilt([charactersToConsider[charactersToConsider.length-1]], [false]);
  displayOutputs(outputs);

  sortedTeamsCheckboxesRange.setValues(oldSortedTeamsCheckboxValues);
  unbuiltCharactersRange.setValues(unbuiltCharacters);
}

function addOutputForCharacter(outputs, character, tierListParams) {
  const metaOutput = calculateCharacterMetaData(tierListParams).get(character);
  outputs.push({
    character : character,
    numMetaTeams : metaOutput.numMetaTeams,
    maxTeamStrength : metaOutput.sortedTeamsStrengths[0],
    tierScore : metaOutput.tierScore,
    tier : metaOutput.tier,
    strongestTeams : metaOutput.strongestTeams,
  });
}

/** Returns the characters that the user checked in the Who To Build sheet. */
function getCharactersToConsider() {
  var charactersToConsider = [];
  for (var r = 0; r < unbuiltCharacters.length; r++) {
    const characterName = unbuiltCharacters[r][0];
    if (characterName === null) {
      return charactersToConsider;
    }
    const considerCharacter = unbuiltCharacters[r][1];
    if (considerCharacter) {
      charactersToConsider.push(characterName);
    }
  }
  return charactersToConsider;
}

function secondsToString(seconds) {
  var minutes = Math.floor((((seconds % 31536000) % 86400) % 3600) / 60);
  if (minutes == 0) minutes = "";
  else if (minutes == 1) minutes += " minute";
  else minutes += " minutes"; 
  var seconds = (((seconds % 31536000) % 86400) % 3600) % 60;
  if (seconds == 0) return minutes;
  return minutes + " " + seconds + " seconds";
}

function roundToNearest15(number) {
  return Math.round(number / 15) * 15;
}