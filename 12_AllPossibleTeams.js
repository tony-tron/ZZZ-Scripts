/** @OnlyCurrentDoc */

var teamCharsToTeamObjs;
var supportedTeamPropertiesToCalcs;

function getAllPossibleTeamsSheet() {
  return getSpreadsheet().getSheetByName("All Possible Teams");
}

function getTeamCharsToTeamObjs() {
  if (!teamCharsToTeamObjs) {
    initializeAllTeamsAndBuffParams();
  }
  return teamCharsToTeamObjs;
}

function getSupportedTeamPropertiesToCalcs() {
  if (!supportedTeamPropertiesToCalcs) {
    initializeAllTeamsAndBuffParams();
  }
  return supportedTeamPropertiesToCalcs;
}

function initializeAllTeamsAndBuffParams() {
  teamCharsToTeamObjs = {};
  supportedTeamPropertiesToCalcs = {};

  // Ensure charsToBuffParams is initialized
  if (typeof getCharsToBuffParams === 'function') {
      getCharsToBuffParams();
  } else if (!charsToBuffParams && typeof initCharsToBuffParams === 'function') {
      charsToBuffParams = initCharsToBuffParams();
  }

  const _teams = getAllPossibleTeamsSheet().getDataRange().getValues();

  for (var r = 0; r < _teams.length; r++) {
    var team = {
      characters : [_teams[r][0], _teams[r][1], _teams[r][2]]
    }
    // Using global charsToBuffParams - assuming it is populated
    if (!charsToBuffParams.has(team.characters[0])) continue;
    addBuffParamsToTeam(team);
    teamCharsToTeamObjs[team.characters.join("|")] = team;
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
  if (typeof getCharsToBuffParams === 'function') {
      return Object.keys(getCharsToBuffParams().get('Anby'));
  }
  // Fallback if not refactored yet
  return Object.keys(charsToBuffParams.get('Anby'));
}
