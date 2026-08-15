/** @OnlyCurrentDoc */

var _teamCharsToTeamObjs;
var supportedTeamPropertiesToCalcs;

var _allPossibleTeamsSheet;

function getAllPossibleTeamsSheet() {
  if (!_allPossibleTeamsSheet) {
    _allPossibleTeamsSheet = getSpreadsheet().getSheetByName("All Possible Teams");
  }
  return _allPossibleTeamsSheet;
}

function getTeamCharsToTeamObjs() {
  if (!_teamCharsToTeamObjs) {
    initializeAllTeamsAndBuffParams();
  }
  return _teamCharsToTeamObjs;
}

function getSupportedTeamPropertiesToCalcs() {
  if (!supportedTeamPropertiesToCalcs) {
    initializeAllTeamsAndBuffParams();
  }
  return supportedTeamPropertiesToCalcs;
}

function initializeAllTeamsAndBuffParams() {
  _teamCharsToTeamObjs = {};
  supportedTeamPropertiesToCalcs = {};

  // Add functions from Team.prototype
  Object.getOwnPropertyNames(Team.prototype).forEach(prop => {
    if (prop !== 'constructor' && prop[0] !== prop[0].toLowerCase()) {
      supportedTeamPropertiesToCalcs[prop] = [];
    }
  });

  // Directly use the getter from 09_BuffUtils.js
  const params = getCharsToBuffParams();

  const _teams = getAllPossibleTeamsSheet().getDataRange().getValues();

  for (var r = 0; r < _teams.length; r++) {
    var char1 = _teams[r][0];
    var char2 = _teams[r][1];
    var char3 = _teams[r][2];

    if (!params.has(char1)) break;

    var team = new Team(char1, char2, char3);

    _teamCharsToTeamObjs[team.characters.join("|")] = team;
    for (const property in team) {
      // By convention, we assume all uppercase properties are variables the user can put in their functions.
      if (property[0] === property[0].toLowerCase()) continue;
      calc = Math.round(Number(team[property]) * 1000) / 1000;
      if (supportedTeamPropertiesToCalcs[property] == undefined) {
        supportedTeamPropertiesToCalcs[property] = [calc];
      } else {
        supportedTeamPropertiesToCalcs[property].push(calc);
      }
    }
  }
}

/**
 * Returns the list of all variables set for each team,
 * which can be used in the buff expressions for
 * Synergy Bonus and Team Synergy. In the second column,
 * provides the min, max, and median values of each property.
 * 
 * @customfunction
 */
function SUPPORTED_TEAM_PROPERTIES() {
  const props = getSupportedTeamPropertiesToCalcs();
  const properties = [];
  var calcs;
  for (const property in props) {
    if (property === "Tags") {
      properties.push([property, "String"]);
      continue;
    }
    if (property === "AnomalyBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "PerAttributeAnomalyBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "DisorderBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "EXSpecialBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "UltimateBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "StunBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "PerChar") {
      properties.push([property, "Function, usage: PerChar('expression')"]);
      continue;
    }
    if (property === "Buff") {
      properties.push([property, "Parameter: attributes"]);
      continue;
    }
    if (property === "Nerf") {
      properties.push([property, "Parameter: attributes"]);
      continue;
    }
    if (!props[property] || props[property].length === 0) {
      properties.push([property, "Function"]);
      continue;
    }
    calcs = props[property].sort((a, b) => a - b);
    properties.push([property, calcs[0] + " to " + calcs[calcs.length - 1] + ", median=" + calcs[Math.floor(calcs.length / 2)]]);
  }
  return properties;
}

/**
 * Returns the list of all variables set for each character,
 * which can be used in the buff expressions for
 * Synergy Bonus and Team Synergy via `PerChar('expression')`.
 * 
 * @customfunction
 */
function SUPPORTED_CHAR_PROPERTIES() {
  return Object.keys(getCharsToBuffParams().get('Anby'));
}

/**
 * Replaces the contents of the AllPossibleTeams sheet with all unique combinations
 * of 3 characters from the Characters sheet, sorted according to specific assist rules.
 */
function populateAllPossibleTeamsSheet() {
  const characters = getCharacterNames();
  const paramsMap = {};
  const buffParams = getCharsToBuffParams();

  for (const char of characters) {
    const params = buffParams.get(char);
    if (params) {
      paramsMap[char] = {
        damageFocus: Number(params.damageFocus || 0),
        quickAssistFocus: Number(params.quickAssistFocus || 0),
        fieldTime: Number(params.fieldTime || 0)
      };
    }
  }

  const validCharacters = characters.filter(c => paramsMap[c]);
  const allTeams = [];

  // Generate all unique combinations of 3 characters
  for (let i = 0; i < validCharacters.length - 2; i++) {
    for (let j = i + 1; j < validCharacters.length - 1; j++) {
      for (let k = j + 1; k < validCharacters.length; k++) {
        const teamChars = [validCharacters[i], validCharacters[j], validCharacters[k]];
        const orderedTeam = _orderTeamForAllPossibleTeams(teamChars, paramsMap);
        allTeams.push(orderedTeam);
      }
    }
  }

  const sheet = getAllPossibleTeamsSheet();
  sheet.clearContents();

  if (allTeams.length > 0) {
    sheet.getRange(1, 1, allTeams.length, 3).setValues(allTeams);
  }
}

/**
 * Helper function to determine the optimal ordering of a 3-character team
 * based on DPS, assist flow, and field time.
 */
function _orderTeamForAllPossibleTeams(teamChars, paramsMap) {
  let dpsChar = teamChars[0];
  for (let i = 1; i < teamChars.length; i++) {
    const candidate = teamChars[i];
    const pDps = paramsMap[dpsChar];
    const pCand = paramsMap[candidate];

    if (Math.abs(pCand.quickAssistFocus) < Math.abs(pDps.quickAssistFocus)) {
      dpsChar = candidate;
    } else if (Math.abs(pCand.quickAssistFocus) === Math.abs(pDps.quickAssistFocus)) {
      if (pCand.damageFocus > pDps.damageFocus) {
        dpsChar = candidate;
      } else if (pCand.damageFocus === pDps.damageFocus) {
        if (pCand.fieldTime > pDps.fieldTime) {
          dpsChar = candidate;
        }
      }
    }
  }

  const getPermutations = (arr) => {
    if (arr.length <= 1) return [arr];
    const result = [];
    for (let i = 0; i < arr.length; i++) {
      const rest = [...arr.slice(0, i), ...arr.slice(i + 1)];
      for (const p of getPermutations(rest)) {
        result.push([arr[i], ...p]);
      }
    }
    return result;
  };

  const permutations = getPermutations(teamChars);
  let bestPerm = null;
  let bestScore = -Infinity;
  let bestPos0FieldTime = Infinity;
  let bestPos2FieldTime = -Infinity;

  const hasRemielle = teamChars.includes('Remielle');

  for (const perm of permutations) {
    if (hasRemielle && perm[0] !== 'Remielle') {
      continue;
    }

    let score = 0;

    // Evaluate assist flow
    for (let i = 0; i < 3; i++) {
      const char = perm[i];
      const p = paramsMap[char];

      let targetIndex = -1;
      if (p.quickAssistFocus > 0) {
        targetIndex = (i + 1) % 3;
      } else if (p.quickAssistFocus < 0) {
        targetIndex = (i + 2) % 3; // Equivalent to (i - 1) wrapped
      }

      if (targetIndex !== -1) {
        const targetChar = perm[targetIndex];
        if (targetChar === dpsChar) {
          score += 100 * Math.abs(p.quickAssistFocus);
        } else {
          const pTarget = paramsMap[targetChar];
          let targetOfTargetIndex = -1;
          if (pTarget.quickAssistFocus > 0) {
            targetOfTargetIndex = (targetIndex + 1) % 3;
          } else if (pTarget.quickAssistFocus < 0) {
            targetOfTargetIndex = (targetIndex + 2) % 3;
          }

          if (targetOfTargetIndex !== -1 && perm[targetOfTargetIndex] === dpsChar) {
            score += 10 * Math.abs(p.quickAssistFocus);
          }
        }
      }
    }

    const pos0FieldTime = paramsMap[perm[0]].fieldTime;
    const pos2FieldTime = paramsMap[perm[2]].fieldTime;

    let isBetter = false;
    if (score > bestScore) {
      isBetter = true;
    } else if (score === bestScore) {
      if (pos0FieldTime < bestPos0FieldTime) {
        isBetter = true;
      } else if (pos0FieldTime === bestPos0FieldTime) {
        if (pos2FieldTime > bestPos2FieldTime) {
          isBetter = true;
        }
      }
    }

    if (isBetter || bestPerm === null) {
      bestScore = score;
      bestPos0FieldTime = pos0FieldTime;
      bestPos2FieldTime = pos2FieldTime;
      bestPerm = perm;
    }
  }

  return bestPerm;
}
