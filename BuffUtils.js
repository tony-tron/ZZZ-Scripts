/** @OnlyCurrentDoc */

var charsToBuffParams = initCharsToBuffParams();
// Cache to store compiled functions. 
// Keys are the formula strings, values are the executable functions.
const formulaCache = {};

// Initializes all of the params per character (aggregating for the team is done in addBuffParamsToTeam).
function initCharsToBuffParams() {
  const charsToBuffParams = new Map();
  const charactersData = charactersDataRange.getValues();
  getCharacterNames().forEach((character, row) => {
    row += 1; // Header.

    // Initialize early so we can reference within the object initialization.
    const specialty = charactersData[row][charactersColumns.specialty];
    const attribute = charactersData[row][charactersColumns.attribute];
    const assistType = charactersData[row][charactersColumns.assistType];
    const anomalyBuildup = Number(charactersData[row][charactersColumns.anomalyBuildup]);
    const fieldTime = Number(charactersData[row][charactersColumns.fieldTime]);
    const damageFocus = Number(charactersData[row][charactersColumns.damageFocus]);
    const anomalyDamage = Number(charactersData[row][charactersColumns.anomalyDamage]);
    const aftershockFocus = Number(charactersData[row][charactersColumns.aftershockFocus]);

    const buffParams = {
      name : character,
      specialty : specialty,
      support : specialty == "Support" ? 1 : 0,
      stun : specialty == "Stun" ? 1 : 0,
      attack : specialty == "Attack" ? 1 : 0,
      anomaly : specialty == "Anomaly" ? 1 : 0,
      defense : specialty == "Defense" ? 1 : 0,
      rupture : specialty == "Rupture" ? 1 : 0,

      attribute : attribute,
      physical : attribute == "Physical" ? 1 : 0,
      ether : attribute == "Ether" ? 1 : 0,
      fire : attribute == "Fire" ? 1 : 0,
      ice : attribute == "Ice" ? 1 : 0,
      electric : attribute == "Electric" ? 1 : 0,

      defensiveAssist : assistType == "Defensive" ? 1 : 0,
      evasiveAssist : assistType == "Evasive" ? 1 : 0,

      tags : String(charactersData[row][charactersColumns.tags]),

      fieldTime : fieldTime,
      stunBuildup : Number(charactersData[row][charactersColumns.stunBuildup]),

      anomalyBuildup : anomalyBuildup,
      physicalAnomalyBuildup : attribute == "Physical" && character != "Ye Shunguang" ? anomalyBuildup : 0,
      honedEdgeAnomalyBuildup : character == "Ye Shunguang" ? anomalyBuildup : 0,
      etherAnomalyBuildup : attribute == "Ether" ? anomalyBuildup : 0,
      fireAnomalyBuildup : attribute == "Fire" ? anomalyBuildup : 0,
      iceAnomalyBuildup : attribute == "Ice" && character != "Miyabi" ? anomalyBuildup : 0,
      frostAnomalyBuildup : character == "Miyabi" ? anomalyBuildup : 0,
      electricAnomalyBuildup : attribute == "Electric" ? anomalyBuildup : 0,

      damageFocus : damageFocus,
      offFieldDamage : fieldTime <= 0 ? damageFocus : 0,
      onFieldDamage : fieldTime > 0 ? damageFocus : 0,
      physicalDamage : attribute == "Physical" ? damageFocus : 0,
      etherDamage : attribute == "Ether" ? damageFocus : 0,
      fireDamage : attribute == "Fire" ? damageFocus : 0,
      iceDamage : attribute == "Ice" ? damageFocus : 0,
      electricDamage : attribute == "Electric" ? damageFocus : 0,
      sheerDamage : specialty == "Rupture" ? damageFocus : 0,
      basicAttack : Number(charactersData[row][charactersColumns.basicAttack]),
      dashAttack : Number(charactersData[row][charactersColumns.dashAttack]),
      dodgeCounter : Number(charactersData[row][charactersColumns.dodgeCounter]),
      assistFollowup : Number(charactersData[row][charactersColumns.assistFollowup]),
      specialAttack : Number(charactersData[row][charactersColumns.specialAttack]),
      exSpecialAttack : Number(charactersData[row][charactersColumns.exSpecialAttack]),
      chainAttack : Number(charactersData[row][charactersColumns.chainAttack]),
      ultimate : Number(charactersData[row][charactersColumns.ultimate]),

      anomalyDamage : anomalyDamage,
      physicalAnomalyDamage : attribute == "Physical" ? anomalyDamage : 0,
      etherAnomalyDamage : attribute == "Ether" ? anomalyDamage : 0,
      fireAnomalyDamage : attribute == "Fire" ? anomalyDamage : 0,
      iceAnomalyDamage : attribute == "Ice" ? anomalyDamage : 0,
      electricAnomalyDamage : attribute == "Electric" ? anomalyDamage : 0,

      shieldFocus : Number(charactersData[row][charactersColumns.shieldFocus]),
      healingFocus : Number(charactersData[row][charactersColumns.healingFocus]),
      quickAssistFocus : Number(charactersData[row][charactersColumns.quickAssistFocus]),

      chainFocus : Number(charactersData[row][charactersColumns.chainFocus]),
      chainEnablement : Number(charactersData[row][charactersColumns.chainEnablement]),
      aftershockFocus : aftershockFocus,
      aftershockDamage : aftershockFocus * damageFocus / 2,
      exSpecialFocus : Number(charactersData[row][charactersColumns.exSpecialFocus]),
      ultimateFocus : Number(charactersData[row][charactersColumns.ultimateFocus]),
      ultimateEnablement : Number(charactersData[row][charactersColumns.ultimateEnablement]),

      hpBenefit : Number(charactersData[row][charactersColumns.hpBenefit]),
      atkBenefit : Number(charactersData[row][charactersColumns.atkBenefit]),
      defBenefit : Number(charactersData[row][charactersColumns.defBenefit]),
      resShredBenefit : Number(charactersData[row][charactersColumns.resShredBenefit]),
      defShredBenefit : Number(charactersData[row][charactersColumns.defShredBenefit]),
      impactBenefit : Number(charactersData[row][charactersColumns.impactBenefit]),
      critRateBenefit : Number(charactersData[row][charactersColumns.critRateBenefit]),
      critDamageBenefit : Number(charactersData[row][charactersColumns.critDamageBenefit]),
      energyRegenBenefit : Number(charactersData[row][charactersColumns.energyRegenBenefit]),
    };

    charsToBuffParams.set(character, buffParams);
  });
  return charsToBuffParams;
}

// Aggregates all of the team params we can reference for buff calculations on the spreadsheet itself.
function addBuffParamsToTeam(team) {
  const char1 = team.characters[0];
  const char2 = team.characters[1];
  const char3 = team.characters[2];
  const char1Params = charsToBuffParams.get(char1);
  const char2Params = charsToBuffParams.get(char2);
  const char3Params = charsToBuffParams.get(char3);

  team.NumSupport = char1Params.support + char2Params.support + char3Params.support;
  team.NumStun = char1Params.stun + char2Params.stun + char3Params.stun;
  team.NumAttack = char1Params.attack + char2Params.attack + char3Params.attack;
  team.NumAnomaly = char1Params.anomaly + char2Params.anomaly + char3Params.anomaly;
  team.NumDefense = char1Params.defense + char2Params.defense + char3Params.defense;
  team.NumRupture = char1Params.rupture + char2Params.rupture + char3Params.rupture;

  team.NumPhysical = char1Params.physical + char2Params.physical + char3Params.physical;
  team.NumEther = char1Params.ether + char2Params.ether + char3Params.ether;
  team.NumFire = char1Params.fire + char2Params.fire + char3Params.fire;
  team.NumIce = char1Params.ice + char2Params.ice + char3Params.ice;
  team.NumElectric = char1Params.electric + char2Params.electric + char3Params.electric;

  team.NumDefensiveAssist = char1Params.defensiveAssist + char2Params.defensiveAssist + char3Params.defensiveAssist;
  team.NumEvasiveAssist = char1Params.evasiveAssist + char2Params.evasiveAssist + char3Params.evasiveAssist;
  team.Tags = char1Params.tags + char2Params.tags + char3Params.tags; // Concatenating strings.
  team.FieldTime = char1Params.fieldTime + char2Params.fieldTime + char3Params.fieldTime;
  team.StunBuildup = char1Params.stunBuildup + char2Params.stunBuildup + char3Params.stunBuildup;

  team.TotalAnomalyBuildup = char1Params.anomalyBuildup + char2Params.anomalyBuildup + char3Params.anomalyBuildup;
  team.PhysicalAnomalyBuildup = char1Params.physicalAnomalyBuildup + char2Params.physicalAnomalyBuildup + char3Params.physicalAnomalyBuildup;
  team.HonedEdgeAnomalyBuildup = char1Params.honedEdgeAnomalyBuildup + char2Params.honedEdgeAnomalyBuildup + char3Params.honedEdgeAnomalyBuildup;
  team.EtherAnomalyBuildup = char1Params.etherAnomalyBuildup + char2Params.etherAnomalyBuildup + char3Params.etherAnomalyBuildup;
  team.FireAnomalyBuildup = char1Params.fireAnomalyBuildup + char2Params.fireAnomalyBuildup + char3Params.fireAnomalyBuildup;
  team.IceAnomalyBuildup = char1Params.iceAnomalyBuildup + char2Params.iceAnomalyBuildup + char3Params.iceAnomalyBuildup;
  team.FrostAnomalyBuildup = char1Params.frostAnomalyBuildup + char2Params.frostAnomalyBuildup + char3Params.frostAnomalyBuildup;
  team.ElectricAnomalyBuildup = char1Params.electricAnomalyBuildup + char2Params.electricAnomalyBuildup + char3Params.electricAnomalyBuildup;

  team.OffFieldDamage = char1Params.offFieldDamage + char2Params.offFieldDamage + char3Params.offFieldDamage;
  team.OnFieldDamage = char1Params.onFieldDamage + char2Params.onFieldDamage + char3Params.onFieldDamage;

  team.TotalDamageFocus = char1Params.damageFocus + char2Params.damageFocus + char3Params.damageFocus;
  team.PhysicalDamage = char1Params.physicalDamage + char2Params.physicalDamage + char3Params.physicalDamage;
  team.EtherDamage = char1Params.etherDamage + char2Params.etherDamage + char3Params.etherDamage;
  team.FireDamage = char1Params.fireDamage + char2Params.fireDamage + char3Params.fireDamage;
  team.IceDamage = char1Params.iceDamage + char2Params.iceDamage + char3Params.iceDamage;
  team.ElectricDamage = char1Params.electricDamage + char2Params.electricDamage + char3Params.electricDamage;
  team.SheerDamage = char1Params.sheerDamage + char2Params.sheerDamage + char3Params.sheerDamage;
  team.BasicAttackDamage = char1Params.basicAttack + char2Params.basicAttack + char3Params.basicAttack;
  team.DashAttackDamage = char1Params.dashAttack + char2Params.dashAttack + char3Params.dashAttack;
  team.DodgeCounterDamage = char1Params.dodgeCounter + char2Params.dodgeCounter + char3Params.dodgeCounter;
  team.AssistFollowupDamage = char1Params.assistFollowup + char2Params.assistFollowup + char3Params.assistFollowup;
  team.SpecialAttackDamage = char1Params.specialAttack + char2Params.specialAttack + char3Params.specialAttack;
  team.EXSpecialAttackDamage = char1Params.exSpecialAttack + char2Params.exSpecialAttack + char3Params.exSpecialAttack;
  team.ChainDamage = char1Params.chainAttack + char2Params.chainAttack + char3Params.chainAttack;
  team.UltimateDamage = char1Params.ultimate + char2Params.ultimate + char3Params.ultimate;

  _updateTeamForYuzuha(team, char1, char2, char3, char1Params, char2Params, char3Params);

  const physicalAnomaly = team.PhysicalAnomalyBuildup >= 2 ? team.PhysicalAnomalyBuildup : 0;
  const honedEdgeAnomaly = team.HonedEdgeAnomalyBuildup >= 2 ? team.HonedEdgeAnomalyBuildup : 0;
  const etherAnomaly = team.EtherAnomalyBuildup >= 2 ? team.EtherAnomalyBuildup : 0;
  const fireAnomaly = team.FireAnomalyBuildup >= 2 ? team.FireAnomalyBuildup : 0;
  const iceAnomaly = team.IceAnomalyBuildup >= 2 ? team.IceAnomalyBuildup : 0;
  const frostAnomaly = team.FrostAnomalyBuildup >= 2 ? team.FrostAnomalyBuildup : 0;
  const electricAnomaly = team.ElectricAnomalyBuildup >= 2 ? team.ElectricAnomalyBuildup : 0;
  const totalAnomaly =
      physicalAnomaly
    + honedEdgeAnomaly
    + etherAnomaly
    + fireAnomaly
    + iceAnomaly
    + frostAnomaly
    + electricAnomaly;
  team.HasAttributeAnomaly = totalAnomaly > 0;
  team.AnomalyBuffUptime = function(uptimeSeconds) {
    with (this) {
      return Math.min(1, totalAnomaly * uptimeSeconds / 60);
    }
  };
  team.DisorderFocus =
      ((team.PhysicalAnomalyBuildup >= 2 ? team.PhysicalAnomalyBuildup : 0)
    + (team.HonedEdgeAnomalyBuildup >= 2 ? team.HonedEdgeAnomalyBuildup : 0)
    + (team.EtherAnomalyBuildup >= 2 ? team.EtherAnomalyBuildup : 0)
    + (team.FireAnomalyBuildup >= 2 ? team.FireAnomalyBuildup : 0)
    + (team.IceAnomalyBuildup >= 2 ? team.IceAnomalyBuildup : 0)
    + (team.FrostAnomalyBuildup >= 2 ? team.FrostAnomalyBuildup : 0)
    + (team.ElectricAnomalyBuildup >= 2 ? team.ElectricAnomalyBuildup : 0))
    / 4;
  if (team.DisorderFocus < 1) team.DisorderFocus = 0;

  team.TotalAnomalyDamage = char1Params.anomalyDamage + char2Params.anomalyDamage + char3Params.anomalyDamage;
  team.PhysicalAnomalyDamage = char1Params.physicalAnomalyDamage + char2Params.physicalAnomalyDamage + char3Params.physicalAnomalyDamage;
  team.EtherAnomalyDamage = char1Params.etherAnomalyDamage + char2Params.etherAnomalyDamage + char3Params.etherAnomalyDamage;
  team.FireAnomalyDamage = char1Params.fireAnomalyDamage + char2Params.fireAnomalyDamage + char3Params.fireAnomalyDamage;
  team.IceAnomalyDamage = char1Params.iceAnomalyDamage + char2Params.iceAnomalyDamage + char3Params.iceAnomalyDamage;
  team.ElectricAnomalyDamage = char1Params.electricAnomalyDamage + char2Params.electricAnomalyDamage + char3Params.electricAnomalyDamage;

  team.ShieldFocus = char1Params.shieldFocus + char2Params.shieldFocus + char3Params.shieldFocus;
  team.HealingFocus = char1Params.healingFocus + char2Params.healingFocus + char3Params.healingFocus;
  team.QuickAssistFocus = Math.abs(char1Params.quickAssistFocus) + Math.abs(char2Params.quickAssistFocus) + Math.abs(char3Params.quickAssistFocus);
  team.ForwardAssistFocus = Math.max(0, char1Params.quickAssistFocus) + Math.max(0, char2Params.quickAssistFocus) + Math.max(0, char3Params.quickAssistFocus);
  team.BackwardAssistFocus = Math.abs(Math.min(0, char1Params.quickAssistFocus) + Math.min(0, char2Params.quickAssistFocus) + Math.min(0, char3Params.quickAssistFocus));

  team.ChainFocus = char1Params.chainFocus + char2Params.chainFocus + char3Params.chainFocus;
  team.ChainEnablement = char1Params.chainEnablement + char2Params.chainEnablement + char3Params.chainEnablement;
  team.AftershockFocus = char1Params.aftershockFocus + char2Params.aftershockFocus + char3Params.aftershockFocus;
  team.EXSpecialFocus = char1Params.exSpecialFocus + char2Params.exSpecialFocus + char3Params.exSpecialFocus;
  team.HasAftershock = team.AftershockFocus > 0;
  team.AftershockDamage = char1Params.aftershockDamage + char2Params.aftershockDamage + char3Params.aftershockDamage;
  team.UltimateFocus = char1Params.ultimateFocus + char2Params.ultimateFocus + char3Params.ultimateFocus;
  team.UltimateEnablement = char1Params.ultimateEnablement + char2Params.ultimateEnablement + char3Params.ultimateEnablement;
  team.EXSpecialBuffUptime = function(uptimeSeconds) {
    with (this) {
      return Math.min(1, EXSpecialFocus * uptimeSeconds / 10);
    }
  };
  team.UltimateBuffUptime = function(uptimeSeconds) {
    with (this) {
      return Math.min(1, UltimateFocus * uptimeSeconds / 60);
    }
  };

  team.HPBenefit = char1Params.hpBenefit + char2Params.hpBenefit + char3Params.hpBenefit;
  team.AttackBenefit = char1Params.atkBenefit + char2Params.atkBenefit + char3Params.atkBenefit;
  team.DefenseBenefit = char1Params.defBenefit + char2Params.defBenefit + char3Params.defBenefit;
  team.ResistanceShredBenefit = char1Params.resShredBenefit + char2Params.resShredBenefit + char3Params.resShredBenefit;
  team.DefenseShredBenefit = char1Params.defShredBenefit + char2Params.defShredBenefit + char3Params.defShredBenefit;
  team.ImpactBenefit = char1Params.impactBenefit + char2Params.impactBenefit + char3Params.impactBenefit;
  team.CritRateBenefit = char1Params.critRateBenefit + char2Params.critRateBenefit + char3Params.critRateBenefit;
  team.CritDamageBenefit = char1Params.critDamageBenefit + char2Params.critDamageBenefit + char3Params.critDamageBenefit;
  team.EnergyRegenBenefit = char1Params.energyRegenBenefit + char2Params.energyRegenBenefit + char3Params.energyRegenBenefit;

  team.StunDamageMultiplier = function(multiplier) {
    return this.PerChar('(name=="Ye Shunguang" ? 1 : ' + (this.StunBuildup*0.1) + ') * damageFocus * ' + multiplier);
  }

  team.PerChar = function(calcExpression) {
    // Handle empty/null expressions gracefully
    if (!calcExpression) return 0;

    if (!formulaCache[calcExpression]) {
      formulaCache[calcExpression] = new Function("ctx", "with(ctx) { return Number(" + calcExpression + "); }");
    }

    const fn = formulaCache[calcExpression];
    return fn(char1Params) + fn(char2Params) + fn(char3Params);
  }

  team.Buff = function(attributes) {
    if (!attributes) return 0;
    
    var total = 0;
    
    // Helper to calculate for one character
    function addForChar(params) {
      if (attributes.toLowerCase().includes(params.attribute.toLowerCase())) {
        total += params.damageFocus*0.2 + params.stunBuildup*0.2 + params.anomalyBuildup*0.2;
      }
    }

    addForChar(char1Params);
    addForChar(char2Params);
    addForChar(char3Params);

    return total;
  };

  team.Nerf = function(attributes) {
    return -this.Buff(attributes);
  };

  team.calculateBuff = function(calcExpression) {
    // Handle empty/null expressions gracefully
    if (!calcExpression) return 0;

    if (!formulaCache[calcExpression]) {
      formulaCache[calcExpression] = new Function("ctx", "with(ctx) { return Number(" + calcExpression + "); }");
    }

    return formulaCache[calcExpression](this);
  };
}

function _updateTeamForYuzuha(team, char1, char2, char3, char1Params, char2Params, char3Params) {
  // Yuzuha matches the attributes applied by the other characters.
  var yuzuhaBuildup = 0;
  var yuzuhaDamage = 0;
  if (char1 == "Yuzuha") {
    yuzuhaBuildup = char1Params.physicalAnomalyBuildup;
    yuzuhaDamage = char1Params.physicalDamage;
  } else if (char2 == "Yuzuha") {
    yuzuhaBuildup = char2Params.physicalAnomalyBuildup;
    yuzuhaDamage = char2Params.physicalDamage;
  } else if (char3 == "Yuzuha") {
    yuzuhaBuildup = char3Params.physicalAnomalyBuildup;
    yuzuhaDamage = char3Params.physicalDamage;
  }

  if (yuzuhaBuildup == 0) return;

  team.PhysicalAnomalyBuildup -= yuzuhaBuildup;
  team.PhysicalAnomalyBuildup += (team.NumPhysical - 1) * yuzuhaBuildup * 0.5;
  team.EtherAnomalyBuildup += team.NumEther * yuzuhaBuildup * 0.5;
  team.FireAnomalyBuildup += team.NumFire * yuzuhaBuildup * 0.5;
  team.IceAnomalyBuildup += team.NumIce * yuzuhaBuildup * 0.5;
  team.ElectricAnomalyBuildup += team.NumElectric * yuzuhaBuildup * 0.5;

  team.PhysicalDamage -= yuzuhaDamage;
  team.PhysicalDamage += (team.NumPhysical - 1) * yuzuhaDamage * 0.5;
  team.EtherDamage += team.NumEther * yuzuhaDamage * 0.5;
  team.FireDamage += team.NumFire * yuzuhaDamage * 0.5;
  team.IceDamage += team.NumIce * yuzuhaDamage * 0.5;
  team.ElectricDamage += team.NumElectric * yuzuhaDamage * 0.5;

  team.NumPhysical -= 1;
  team.NumPhysical *= 1.5;
  team.NumEther *= 1.5;
  team.NumFire *= 1.5;
  team.NumIce *= 1.5;
  team.NumElectric *= 1.5;
}

/**
 * Calculates the total strength bonus for a single team from a list of buff expressions.
 * @param {object} team - A team object with a calculateBuff method.
 * @param {Array<string>} teamBuffExpressions - A list of buff expressions to apply.
 * @returns {number} The total calculated buff strength.
 */
function computeStrengthFromTeamBuffs(team, teamBuffExpressions) {
  var addedStrength = 0;
  teamBuffExpressions.forEach(buffCalc => addedStrength += team.calculateBuff(buffCalc));
  return Math.round(addedStrength * 1000) / 1000;
}

/**
 * Returns the name of the buff that computed the highest bonus, and that computed bonus.
 * @param {object} team - A team object with a calculateBuff method.
 * @param {object} buffOptions - A list of buff formulas that can be chosen for the team.
 * @returns {object | undefined} An object { name, bonus } or undefined.
 */
function chooseBestBuffForTeam(team, buffOptions) {
  if (!buffOptions || buffOptions.length === 0) {
    return;
  }

  var bestBuffBonus = 0;
  var bestBuffOption;

  buffOptions.forEach(buffOption => {
    const buffBonus = team.calculateBuff(buffOption.expression);
    if (buffBonus > bestBuffBonus) {
      bestBuffBonus = buffBonus;
      bestBuffOption = buffOption;
    }
  });

  return {
    name: bestBuffOption.name,
    bonus: Math.round(bestBuffBonus * 1000) / 1000,
  };
}

// --- NEW GENERIC OPTIMIZER ---

/**
 * Finds the optimal assignment of teams to buff slots to maximize total strength.
 *
 * @param {Array<Object>} teamSlots - An array of "team slot" objects. Each object should
 * contain the team and any other data that should
 * be re-ordered with it (e.g., bonus names).
 * Example: [{ team: teamA, name: 'Slot 1' }, { team: teamB, name: 'Slot 2' }]
 * @param {Array<Array<string>>} buffExpressionsList - An array of buff expression arrays,
 * one for each "slot".
 * Example: [ ['buffA'], ['buffB'] ]
 * @returns {Object} An object { optimizedSlots, bonuses, maxStrength }
 * - optimizedSlots: The re-ordered array of teamSlots for max strength.
 * - bonuses: An array of calculated bonuses for each slot.
 * - maxStrength: The total maximized strength.
 */
function optimizeTeamAssignments(teamSlots, buffExpressionsList) {
  const n = teamSlots.length;

  // --- Input Validation ---
  if (n === 0) {
    console.warn("No teams provided to optimize.");
    return { optimizedSlots: [], bonuses: [], maxStrength: 0 };
  }
  if (!buffExpressionsList || n !== buffExpressionsList.length) {
    console.error("Team count and buff list count must match.");
    return { optimizedSlots: teamSlots, bonuses: Array(n).fill(0), maxStrength: 0 };
  }
  // 8! = 40,320 (fast). 9! = 362,880 (getting slow). 10! = 3.6M (too slow for JS).
  // User asked for 5, so 8 is a safe and generous upper limit for this algorithm.
  if (n > 8) {
    console.error(`Cannot optimize more than 8 teams with this permutation method (received ${n}).`);
    return { optimizedSlots: teamSlots, bonuses: Array(n).fill(0), maxStrength: 0 };
  }

  // --- 1. Compute N x N Strength Matrix ---
  // strengthMatrix[i][j] = strength of team `i` with buff slot `j`
  const strengthMatrix = [];
  for (let i = 0; i < n; i++) { // i = team index
    strengthMatrix[i] = [];
    for (let j = 0; j < n; j++) { // j = buff slot index
      const team = teamSlots[i] ? teamSlots[i].team : null;
      strengthMatrix[i][j] = computeStrengthFromTeamBuffs(team, buffExpressionsList[j]);
    }
  }

  // --- 2. Get all permutations of team indices ---
  const indices = Array.from({ length: n }, (_, i) => i);
  const permutations = getAllPermutations(indices);

  // --- 3. Find Best Permutation ---
  let maxStrength = -Infinity;
  let bestPermutation = indices; // Default to no change

  // Iterate over each possible assignment permutation
  for (const perm of permutations) {
    let currentStrength = 0;
    // perm is an array of team indices, e.g., [1, 0, 2]
    // This means team 1 goes to slot 0, team 0 to slot 1, team 2 to slot 2
    for (let j = 0; j < n; j++) { // j is the *slot index*
      const teamIndex = perm[j]; // team index assigned to slot j
      currentStrength += strengthMatrix[teamIndex][j];
    }

    if (currentStrength > maxStrength) {
      maxStrength = currentStrength;
      bestPermutation = perm;
    }
  }

  // --- 4. Build and Return Result ---
  
  // Create the new, optimized array of team slots
  const optimizedSlots = bestPermutation.map(teamIndex => teamSlots[teamIndex]);

  // Create the corresponding array of bonuses
  const bonuses = bestPermutation.map((teamIndex, j) => {
    // teamIndex is the team, j is the slot it was assigned to
    return strengthMatrix[teamIndex][j];
  });

  return {
    optimizedSlots,
    bonuses,
    maxStrength
  };
}

/**
 * Helper function to generate all permutations of an array.
 * @param {Array<any>} arr - The array to permute (e.g., [0, 1, 2])
 * @returns {Array<Array<any>>} An array of all possible permutations.
 */
function getAllPermutations(arr) {
  const result = [];

  function permute(currentArr, remainingArr) {
    if (remainingArr.length === 0) {
      result.push(currentArr);
      return;
    }

    for (let i = 0; i < remainingArr.length; i++) {
      const newCurrent = currentArr.concat(remainingArr[i]);
      // Create new remaining array by removing the i-th element
      const newRemaining = remainingArr.slice(0, i).concat(remainingArr.slice(i + 1));
      permute(newCurrent, newRemaining);
    }
  }

  permute([], arr);
  return result;
}

/**
 * Evaluates the javascript buff function based on the columns:
 * Character1 Character2 Character3 buffFunction
 * 
 * @customfunction
 */
function CALCULATE_BUFFS(charactersAndBuffExpressions) {
  const buffs = [];
  for (var r = 0; r < charactersAndBuffExpressions.length; r++) {
    var character1 = charactersAndBuffExpressions[r][0];
    if (character1 == null || character1 == "") break;
    var character2 = charactersAndBuffExpressions[r][1];
    var character3 = charactersAndBuffExpressions[r][2];
    var buffExpression = charactersAndBuffExpressions[r][3];
    var buff = teamCharsToTeamObjs[[character1, character2, character3].join("|")].calculateBuff(buffExpression);
    // Round to the nearest 0.25.
    buff = Math.round(buff * 4) / 4;
    buffs.push([buff]);
  }
  return buffs;
}
/**
 * Evaluates the javascript buff function based on the columns:
 * Character1 Character2 Character3 char1HasSynergy char2HasSynergy char3HasSynergy char1BuffFunction char2BuffFunction char3BuffFunction
 * * (Synergy, as always, means the given character's Additional Ability is activated.)
 * * @customfunction
 */
function CALCULATE_SYNERGY_BUFFS(data) {
  const results = [];
  
  // Create a local reference to the global object to ensure scope access within the loop
  const lookupMap = teamCharsToTeamObjs; 

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const char1 = row[0];

    // Optimization: Break immediately on empty rows. 
    // This is crucial for Sheets inputs like "A2:I" which pass 1000 rows even if 900 are empty.
    if (!char1) break;

    const char2 = row[1];
    const char3 = row[2];

    // Optimization: Construct the key once per row instead of 3 times.
    // Template literals are generally faster/cleaner than array joins for fixed sizes.
    const key = `${char1}|${char2}|${char3}`;
    const teamObj = lookupMap[key];

    let totalBuff = 0;

    // Direct index access is faster than destructuring inside a hot loop.
    // Indexes: 3/4/5 are Synergy Booleans, 6/7/8 are Buff Expressions.
    if (row[3]) totalBuff += teamObj.calculateBuff(row[6]);
    if (row[4]) totalBuff += teamObj.calculateBuff(row[7]);
    if (row[5]) totalBuff += teamObj.calculateBuff(row[8]);

    // Round to the nearest 0.25
    results.push([Math.round(totalBuff * 4) / 4]);
  }

  return results;
}
