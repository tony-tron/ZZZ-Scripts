/** @OnlyCurrentDoc */

// Refactored to remove global spreadsheet reads.

var charactersColumns; // Lazy loaded
var _characterNames2D; // Cache for character names
var _charactersSheet;
var _characterHeadersRange;
var _characterNamesRange;
var _characterBuiltRange;
var _charactersDataRange;

function getCharactersSheet() {
  if (!_charactersSheet) {
    _charactersSheet = getSpreadsheet().getSheetByName("Characters");
  }
  return _charactersSheet;
}

function getCharacterHeadersRange() {
  if (!_characterHeadersRange) {
    _characterHeadersRange = getCharactersSheet().getRange("A1:ZZZ1");
  }
  return _characterHeadersRange;
}

function getCharacterNamesRange() {
  if (!_characterNamesRange) {
    _characterNamesRange = getCharactersSheet().getRange("A2:A");
  }
  return _characterNamesRange;
}

function getCharacterNames2D() {
  if (!_characterNames2D) {
    _characterNames2D = getCharacterNamesRange().getValues();
  }
  return _characterNames2D;
}

function getCharacterBuiltRange() {
  if (!_characterBuiltRange) {
    _characterBuiltRange = getCharactersSheet().getRange("B2:B");
  }
  return _characterBuiltRange;
}

function getCharactersDataRange() {
  if (!_charactersDataRange) {
    _charactersDataRange = getCharactersSheet().getDataRange();
  }
  return _charactersDataRange;
}

function getCharactersColumns() {
  if (!charactersColumns) {
    charactersColumns = initCharactersColumns();
  }
  return charactersColumns;
}

function initCharactersColumns() {
  const characterColumns = {
    character : 0,
    built : 1,
    additionalAbilityQuery : 2,
    specialty : 3,
    attribute : 4,
    coreStat : 5,
    attackType : 6,
    faction : 7,
    assistType : 8,
    baseStrength : 9,
    synergyBonus : 10,
    extraSynergy : 11,
    extraSynergyBonus : 12,
    teamSynergy : 13,
    tags : 19,
    fieldTime : 20,
    stunBuildup : 21,
    anomalyBuildup : 22,
    damageFocus : 23,
    basicAttack : 24,
    dashAttack : 25,
    dodgeCounter : 26,
    assistFollowup : 27,
    specialAttack : 28,
    exSpecialAttack : 29,
    chainAttack : 30,
    ultimate : 31,
    anomalyDamage : 32,
    otherDamage : 33,
    shieldFocus : 34,
    healingFocus : 35,
    quickAssistFocus : 36,
    chainFocus : 37,
    chainEnablement : 38,
    aftershockFocus : 39,
    exSpecialFocus : 40,
    ultimateFocus : 41,
    ultimateEnablement : 42,
    hpBenefit : 43,
    atkBenefit : 44,
    defBenefit : 45,
    resShredBenefit : 46,
    defShredBenefit : 47,
    impactBenefit : 48,
    critRateBenefit : 49,
    critDamageBenefit : 50,
    energyRegenBenefit : 51,
  };
  return characterColumns;
}

function setCharactersBuilt(characterNames, builts) {
  const builtCharactersRange = getCharacterBuiltRange();
  const builtCharacters = builtCharactersRange.getValues();
  for (var i = 0; i < characterNames.length; i++) {
    const characterName = characterNames[i];
    const built = builts[i];
    const characterRowIndex = getCharacterRowIndex(characterName);
    if (characterRowIndex < 0) {
      break;
    }
    builtCharacters[characterRowIndex][0] = built;
  }
  getCharactersSheet().getRange(builtCharactersRange.getRow(),
    builtCharactersRange.getColumn(), builtCharactersRange.getNumRows())
    .setValues(builtCharacters);
}

function getCharacterNames() {
  const characterNames = [];
  const names2D = getCharacterNames2D();
  for (var r = 0; r < names2D.length; r++) {
    const characterName = names2D[r][0];
    if (characterName == null || characterName == "") break;
    characterNames.push(characterName);
  }
  return characterNames;
}

function getCharacterRowIndex(name) {
  return getCharacterNames().findIndex(charName => charName === name);
}

function getNumCharacters() {
  return getCharacterNames().length;
}

/**
 * Translates a partial query string using header names (in the Characters sheet) into one using column letters.
 *
 * @param {string} partialQueryWithNames The partial query string using header names (e.g., "FieldTime > 0").
 * @return The translated partial query string with column letters (e.g., "U > 0") or an error string.
 * @customfunction
 */
function QUERY_VARIABLE_NAMES_TO_COLUMNS(partialQueryWithNames) {
  // Input validation
  if (typeof partialQueryWithNames !== 'string') {
    return '#VALUE! Invalid input query string.';
  }

  try {
    // 1. Get Headers and Starting Column
    const headersRange = getCharacterHeadersRange();
    const headers = headersRange.getValues()[0];
    const firstColIndex = headersRange.getColumn(); // 1-based index

    // 2. Create Mapping from Header Name to Column Letter
    const nameToLetterMap = {};
    headers.forEach((header, index) => {
      if (header && typeof header === 'string' && header.trim() !== '') {
        const currentColumnIndex = firstColIndex + index; // Calculate absolute column index
        nameToLetterMap[header.replaceAll(" ", "").trim()] = getColumnLetter(currentColumnIndex);
      }
    });

    // Check if map is empty (maybe bad range or empty headers)
    if (Object.keys(nameToLetterMap).length === 0) {
        return "#REF! Could not read headers or headers are empty.";
    }

    // 3. Translate Query String
    let translatedQuery = partialQueryWithNames;
    const foundNames = new Set(); // Keep track of names found in the query

    // Sort names by length descending to handle partial matches (e.g., "Timestamp" before "Time")
    const sortedNames = Object.keys(nameToLetterMap).sort((a, b) => b.length - a.length);

    sortedNames.forEach(name => {
      // Use regex to replace whole words matching the header name (case-sensitive)
      // Escape special regex characters in the header name
      const escapedName = name.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      const regex = new RegExp(`\\b${escapedName}\\b`, 'g');

      // Check if the name exists in the query before replacing
      if (regex.test(translatedQuery)) {
          foundNames.add(name); // Mark this name as found
          translatedQuery = translatedQuery.replace(regex, nameToLetterMap[name]);
      }
    });


    return translatedQuery;

  } catch (e) {
    Logger.log(`QUERY_HELPER Error: ${e}`);
    return `#ERROR! ${e.message}`;
  }
}

/**
 * Converts a column index (1-based) to a spreadsheet column letter (A, B, ..., Z, AA, ...).
 */
function getColumnLetter(columnIndex) {
  let letter = '';
  let temp;
  while (columnIndex > 0) {
    temp = (columnIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnIndex = (columnIndex - temp - 1) / 26;
  }
  return letter;
}
