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
    if (property === "UltimateBuffUptime") {
      properties.push([property, "Parameter: uptimeSeconds"]);
      continue;
    }
    if (property === "StunDamageMultiplier") {
      properties.push([property, "Function"]);
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
