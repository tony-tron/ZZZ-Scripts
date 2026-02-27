/** @OnlyCurrentDoc */

var _charsToBuffParams;

function getCharsToBuffParams() {
  if (!_charsToBuffParams) {
    _charsToBuffParams = initCharsToBuffParams();
  }
  return _charsToBuffParams;
}

// Cache to store compiled functions. 
// Keys are the formula strings, values are the executable functions.
const formulaCache = {};

// Initializes all of the params per character (aggregating for the team is done in addBuffParamsToTeam).
function initCharsToBuffParams() {
  const charsToBuffParams = new Map();
  const charactersData = getCharactersDataRange().getValues();
  const cols = getCharactersColumns();

  getCharacterNames().forEach((character, row) => {
    row += 1; // Header.

    // Initialize early so we can reference within the object initialization.
    const specialty = charactersData[row][cols.specialty];
    const attribute = charactersData[row][cols.attribute];
    const assistType = charactersData[row][cols.assistType];
    const anomalyBuildup = Number(charactersData[row][cols.anomalyBuildup]);
    const fieldTime = Number(charactersData[row][cols.fieldTime]);
    const damageFocus = Number(charactersData[row][cols.damageFocus]);
    const anomalyDamage = Number(charactersData[row][cols.anomalyDamage]);
    const aftershockFocus = Number(charactersData[row][cols.aftershockFocus]);

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

      teamBuffFunction : compileBuffFunction(String(charactersData[row][cols.teamBuffFormula] || "0")),
      tags : String(charactersData[row][cols.tags]),

      fieldTime : fieldTime,
      stunBuildup : Number(charactersData[row][cols.stunBuildup]),

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
      basicAttack : Number(charactersData[row][cols.basicAttack]),
      dashAttack : Number(charactersData[row][cols.dashAttack]),
      dodgeCounter : Number(charactersData[row][cols.dodgeCounter]),
      assistFollowup : Number(charactersData[row][cols.assistFollowup]),
      specialAttack : Number(charactersData[row][cols.specialAttack]),
      exSpecialAttack : Number(charactersData[row][cols.exSpecialAttack]),
      chainAttack : Number(charactersData[row][cols.chainAttack]),
      ultimate : Number(charactersData[row][cols.ultimate]),

      anomalyDamage : anomalyDamage,
      physicalAnomalyDamage : attribute == "Physical" ? anomalyDamage : 0,
      etherAnomalyDamage : attribute == "Ether" ? anomalyDamage : 0,
      fireAnomalyDamage : attribute == "Fire" ? anomalyDamage : 0,
      iceAnomalyDamage : attribute == "Ice" ? anomalyDamage : 0,
      electricAnomalyDamage : attribute == "Electric" ? anomalyDamage : 0,

      shieldFocus : Number(charactersData[row][cols.shieldFocus]),
      healingFocus : Number(charactersData[row][cols.healingFocus]),
      quickAssistFocus : Number(charactersData[row][cols.quickAssistFocus]),

      chainFocus : Number(charactersData[row][cols.chainFocus]),
      chainEnablement : Number(charactersData[row][cols.chainEnablement]),
      aftershockFocus : aftershockFocus,
      aftershockDamage : aftershockFocus * damageFocus / 2,
      exSpecialFocus : Number(charactersData[row][cols.exSpecialFocus]),
      exSpecialBuffUptime: function(uptimeSeconds) {
        return Math.min(1, this.exSpecialFocus * uptimeSeconds / 10);
      },
      ultimateFocus : Number(charactersData[row][cols.ultimateFocus]),
      ultimateBuffUptime: function(uptimeSeconds) {
        return Math.min(1, this.ultimateFocus * uptimeSeconds / 60);
      },
      ultimateEnablement : Number(charactersData[row][cols.ultimateEnablement]),

      hpBenefit : Number(charactersData[row][cols.hpBenefit]),
      atkBenefit : Number(charactersData[row][cols.atkBenefit]),
      defBenefit : Number(charactersData[row][cols.defBenefit]),
      resShredBenefit : Number(charactersData[row][cols.resShredBenefit]),
      defShredBenefit : Number(charactersData[row][cols.defShredBenefit]),
      impactBenefit : Number(charactersData[row][cols.impactBenefit]),
      critRateBenefit : Number(charactersData[row][cols.critRateBenefit]),
      critDamageBenefit : Number(charactersData[row][cols.critDamageBenefit]),
      energyRegenBenefit : Number(charactersData[row][cols.energyRegenBenefit]),
      etherVeilFocus : Number(charactersData[row][cols.etherVeilFocus]),
    };

    charsToBuffParams.set(character, buffParams);
  });
  return charsToBuffParams;
}

// Aggregates all of the team params we can reference for buff calculations on the spreadsheet itself.
class Team {
  constructor(char1, char2, char3) {
    this.characters = [char1, char2, char3];
    this.initStats();
  }

  initStats() {
    const paramsMap = getCharsToBuffParams();
    const p1 = paramsMap.get(this.characters[0]);
    const p2 = paramsMap.get(this.characters[1]);
    const p3 = paramsMap.get(this.characters[2]);

    if (!p1 || !p2 || !p3) return;

    this.NumSupport = p1.support + p2.support + p3.support;
    this.NumStun = p1.stun + p2.stun + p3.stun;
    this.NumAttack = p1.attack + p2.attack + p3.attack;
    this.NumAnomaly = p1.anomaly + p2.anomaly + p3.anomaly;
    this.NumDefense = p1.defense + p2.defense + p3.defense;
    this.NumRupture = p1.rupture + p2.rupture + p3.rupture;

    this.NumPhysical = p1.physical + p2.physical + p3.physical;
    this.NumEther = p1.ether + p2.ether + p3.ether;
    this.NumFire = p1.fire + p2.fire + p3.fire;
    this.NumIce = p1.ice + p2.ice + p3.ice;
    this.NumElectric = p1.electric + p2.electric + p3.electric;

    this.NumDefensiveAssist = p1.defensiveAssist + p2.defensiveAssist + p3.defensiveAssist;
    this.NumEvasiveAssist = p1.evasiveAssist + p2.evasiveAssist + p3.evasiveAssist;
    this.Tags = p1.tags + p2.tags + p3.tags;
    this.FieldTime = p1.fieldTime + p2.fieldTime + p3.fieldTime;
    this.StunBuildup = p1.stunBuildup + p2.stunBuildup + p3.stunBuildup;

    this.TotalAnomalyBuildup = p1.anomalyBuildup + p2.anomalyBuildup + p3.anomalyBuildup;
    this.PhysicalAnomalyBuildup = p1.physicalAnomalyBuildup + p2.physicalAnomalyBuildup + p3.physicalAnomalyBuildup;
    this.HonedEdgeAnomalyBuildup = p1.honedEdgeAnomalyBuildup + p2.honedEdgeAnomalyBuildup + p3.honedEdgeAnomalyBuildup;
    this.EtherAnomalyBuildup = p1.etherAnomalyBuildup + p2.etherAnomalyBuildup + p3.etherAnomalyBuildup;
    this.FireAnomalyBuildup = p1.fireAnomalyBuildup + p2.fireAnomalyBuildup + p3.fireAnomalyBuildup;
    this.IceAnomalyBuildup = p1.iceAnomalyBuildup + p2.iceAnomalyBuildup + p3.iceAnomalyBuildup;
    this.FrostAnomalyBuildup = p1.frostAnomalyBuildup + p2.frostAnomalyBuildup + p3.frostAnomalyBuildup;
    this.ElectricAnomalyBuildup = p1.electricAnomalyBuildup + p2.electricAnomalyBuildup + p3.electricAnomalyBuildup;

    this.OffFieldDamage = p1.offFieldDamage + p2.offFieldDamage + p3.offFieldDamage;
    this.OnFieldDamage = p1.onFieldDamage + p2.onFieldDamage + p3.onFieldDamage;

    this.TotalDamageFocus = p1.damageFocus + p2.damageFocus + p3.damageFocus;
    this.PhysicalDamage = p1.physicalDamage + p2.physicalDamage + p3.physicalDamage;
    this.EtherDamage = p1.etherDamage + p2.etherDamage + p3.etherDamage;
    this.FireDamage = p1.fireDamage + p2.fireDamage + p3.fireDamage;
    this.IceDamage = p1.iceDamage + p2.iceDamage + p3.iceDamage;
    this.ElectricDamage = p1.electricDamage + p2.electricDamage + p3.electricDamage;
    this.SheerDamage = p1.sheerDamage + p2.sheerDamage + p3.sheerDamage;
    this.BasicAttackDamage = p1.basicAttack + p2.basicAttack + p3.basicAttack;
    this.DashAttackDamage = p1.dashAttack + p2.dashAttack + p3.dashAttack;
    this.DodgeCounterDamage = p1.dodgeCounter + p2.dodgeCounter + p3.dodgeCounter;
    this.AssistFollowupDamage = p1.assistFollowup + p2.assistFollowup + p3.assistFollowup;
    this.SpecialAttackDamage = p1.specialAttack + p2.specialAttack + p3.specialAttack;
    this.EXSpecialAttackDamage = p1.exSpecialAttack + p2.exSpecialAttack + p3.exSpecialAttack;
    this.ChainDamage = p1.chainAttack + p2.chainAttack + p3.chainAttack;
    this.UltimateDamage = p1.ultimate + p2.ultimate + p3.ultimate;

    _updateTeamForSunna(this, p1.name, p2.name, p3.name, p1, p2, p3);
    _updateTeamForYuzuha(this, p1.name, p2.name, p3.name, p1, p2, p3);

    const physicalAnomaly = this.PhysicalAnomalyBuildup >= 2 ? this.PhysicalAnomalyBuildup : 0;
    const honedEdgeAnomaly = this.HonedEdgeAnomalyBuildup >= 2 ? this.HonedEdgeAnomalyBuildup : 0;
    const etherAnomaly = this.EtherAnomalyBuildup >= 2 ? this.EtherAnomalyBuildup : 0;
    const fireAnomaly = this.FireAnomalyBuildup >= 2 ? this.FireAnomalyBuildup : 0;
    const iceAnomaly = this.IceAnomalyBuildup >= 2 ? this.IceAnomalyBuildup : 0;
    const frostAnomaly = this.FrostAnomalyBuildup >= 2 ? this.FrostAnomalyBuildup : 0;
    const electricAnomaly = this.ElectricAnomalyBuildup >= 2 ? this.ElectricAnomalyBuildup : 0;
    const totalAnomaly =
        physicalAnomaly
      + honedEdgeAnomaly
      + etherAnomaly
      + fireAnomaly
      + iceAnomaly
      + frostAnomaly
      + electricAnomaly;

    this.HasAttributeAnomaly = totalAnomaly > 0;
    this.totalAnomaly = totalAnomaly;

    this.DisorderFocus =
        ((this.PhysicalAnomalyBuildup >= 2 ? this.PhysicalAnomalyBuildup : 0)
      + (this.HonedEdgeAnomalyBuildup >= 2 ? this.HonedEdgeAnomalyBuildup : 0)
      + (this.EtherAnomalyBuildup >= 2 ? this.EtherAnomalyBuildup : 0)
      + (this.FireAnomalyBuildup >= 2 ? this.FireAnomalyBuildup : 0)
      + (this.IceAnomalyBuildup >= 2 ? this.IceAnomalyBuildup : 0)
      + (this.FrostAnomalyBuildup >= 2 ? this.FrostAnomalyBuildup : 0)
      + (this.ElectricAnomalyBuildup >= 2 ? this.ElectricAnomalyBuildup : 0))
      / 4;
    if (this.DisorderFocus < 1) this.DisorderFocus = 0;

    this.TotalAnomalyDamage = p1.anomalyDamage + p2.anomalyDamage + p3.anomalyDamage;
    this.PhysicalAnomalyDamage = p1.physicalAnomalyDamage + p2.physicalAnomalyDamage + p3.physicalAnomalyDamage;
    this.EtherAnomalyDamage = p1.etherAnomalyDamage + p2.etherAnomalyDamage + p3.etherAnomalyDamage;
    this.FireAnomalyDamage = p1.fireAnomalyDamage + p2.fireAnomalyDamage + p3.fireAnomalyDamage;
    this.IceAnomalyDamage = p1.iceAnomalyDamage + p2.iceAnomalyDamage + p3.iceAnomalyDamage;
    this.ElectricAnomalyDamage = p1.electricAnomalyDamage + p2.electricAnomalyDamage + p3.electricAnomalyDamage;

    this.ShieldFocus = p1.shieldFocus + p2.shieldFocus + p3.shieldFocus;
    this.HealingFocus = p1.healingFocus + p2.healingFocus + p3.healingFocus;
    this.QuickAssistFocus = Math.abs(p1.quickAssistFocus) + Math.abs(p2.quickAssistFocus) + Math.abs(p3.quickAssistFocus);
    this.ForwardAssistFocus = Math.max(0, p1.quickAssistFocus) + Math.max(0, p2.quickAssistFocus) + Math.max(0, p3.quickAssistFocus);
    this.BackwardAssistFocus = Math.abs(Math.min(0, p1.quickAssistFocus) + Math.min(0, p2.quickAssistFocus) + Math.min(0, p3.quickAssistFocus));

    this.ChainFocus = p1.chainFocus + p2.chainFocus + p3.chainFocus;
    this.ChainEnablement = p1.chainEnablement + p2.chainEnablement + p3.chainEnablement;
    this.AftershockFocus = p1.aftershockFocus + p2.aftershockFocus + p3.aftershockFocus;
    this.EXSpecialFocus = p1.exSpecialFocus + p2.exSpecialFocus + p3.exSpecialFocus;
    this.HasAftershock = this.AftershockFocus > 0;
    this.AftershockDamage = p1.aftershockDamage + p2.aftershockDamage + p3.aftershockDamage;
    this.UltimateFocus = p1.ultimateFocus + p2.ultimateFocus + p3.ultimateFocus;
    this.UltimateEnablement = p1.ultimateEnablement + p2.ultimateEnablement + p3.ultimateEnablement;

    this.HPBenefit = p1.hpBenefit + p2.hpBenefit + p3.hpBenefit;
    this.AttackBenefit = p1.atkBenefit + p2.atkBenefit + p3.atkBenefit;
    this.DefenseBenefit = p1.defBenefit + p2.defBenefit + p3.defBenefit;
    this.ResistanceShredBenefit = p1.resShredBenefit + p2.resShredBenefit + p3.resShredBenefit;
    this.DefenseShredBenefit = p1.defShredBenefit + p2.defShredBenefit + p3.defShredBenefit;
    this.ImpactBenefit = p1.impactBenefit + p2.impactBenefit + p3.impactBenefit;
    this.CritRateBenefit = p1.critRateBenefit + p2.critRateBenefit + p3.critRateBenefit;
    this.CritDamageBenefit = p1.critDamageBenefit + p2.critDamageBenefit + p3.critDamageBenefit;
    this.EnergyRegenBenefit = p1.energyRegenBenefit + p2.energyRegenBenefit + p3.energyRegenBenefit;

    this.EtherVeilFocus = p1.etherVeilFocus + p2.etherVeilFocus + p3.etherVeilFocus;
  }

  AnomalyBuffUptime(uptimeSeconds) {
    return Math.min(1, (this.totalAnomaly || 0) * uptimeSeconds / 60);
  }

  EXSpecialBuffUptime(uptimeSeconds) {
      return Math.min(1, this.EXSpecialFocus * uptimeSeconds / 10);
  }

  UltimateBuffUptime(uptimeSeconds) {
      return Math.min(1, this.UltimateFocus * uptimeSeconds / 60);
  }

  StunDamageMultiplier(multiplier) {
    return this.PerChar('(name=="Ye Shunguang" ? 1 : ' + (this.StunBuildup*0.1) + ') * damageFocus * ' + multiplier);
  }

  PerChar(calcExpression) {
    if (!calcExpression) return 0;

    if (!formulaCache[calcExpression]) {
      formulaCache[calcExpression] = new Function("ctx", "with(ctx) { return Number(" + calcExpression + "); }");
    }

    const fn = formulaCache[calcExpression];
    const params = getCharsToBuffParams();
    const p1 = params.get(this.characters[0]);
    const p2 = params.get(this.characters[1]);
    const p3 = params.get(this.characters[2]);

    return fn(p1) + fn(p2) + fn(p3);
  }

  Buff(attributes) {
    if (!attributes) return 0;
    
    var total = 0;
    
    const params = getCharsToBuffParams();
    const p1 = params.get(this.characters[0]);
    const p2 = params.get(this.characters[1]);
    const p3 = params.get(this.characters[2]);

    function addForChar(p) {
      if (attributes.toLowerCase().includes(p.attribute.toLowerCase())) {
        total += p.damageFocus*0.2 + p.stunBuildup*0.2 + p.anomalyBuildup*0.2;
      }
    }

    addForChar(p1);
    addForChar(p2);
    addForChar(p3);

    return total;
  }

  Nerf(attributes) {
    return -this.Buff(attributes);
  }

  calculateBuff(calcExpression) {
    if (!calcExpression) return 0;

    if (!formulaCache[calcExpression]) {
      formulaCache[calcExpression] = new Function("ctx", "with(ctx) { return Number(" + calcExpression + "); }");
    }

    return formulaCache[calcExpression](this);
  }
}

function _updateTeamForSunna(team, char1, char2, char3, char1Params, char2Params, char3Params) {
  var sunnaParams = null;
  var otherParams = [];

  if (char1 == "Sunna") {
    sunnaParams = char1Params;
    otherParams.push(char2Params, char3Params);
  } else if (char2 == "Sunna") {
    sunnaParams = char2Params;
    otherParams.push(char1Params, char3Params);
  } else if (char3 == "Sunna") {
    sunnaParams = char3Params;
    otherParams.push(char1Params, char2Params);
  }

  if (!sunnaParams) return;

  var sunnaBuildup = sunnaParams.anomalyBuildup;
  var sunnaDamage = sunnaParams.damageFocus;
  var sourceAttribute = sunnaParams.attribute;

  // Remove from Source
  if (sourceAttribute == "Physical") {
    team.PhysicalAnomalyBuildup -= sunnaBuildup;
    team.PhysicalDamage -= sunnaDamage;
    team.NumPhysical -= 1;
  } else if (sourceAttribute == "Ether") {
    team.EtherAnomalyBuildup -= sunnaBuildup;
    team.EtherDamage -= sunnaDamage;
    team.NumEther -= 1;
  } else if (sourceAttribute == "Fire") {
    team.FireAnomalyBuildup -= sunnaBuildup;
    team.FireDamage -= sunnaDamage;
    team.NumFire -= 1;
  } else if (sourceAttribute == "Ice") {
    team.IceAnomalyBuildup -= sunnaBuildup;
    team.IceDamage -= sunnaDamage;
    team.NumIce -= 1;
  } else if (sourceAttribute == "Electric") {
    team.ElectricAnomalyBuildup -= sunnaBuildup;
    team.ElectricDamage -= sunnaDamage;
    team.NumElectric -= 1;
  }

  // Calculate qualifying damage focus for proportional split
  var totalQualifyingDamageFocus = 0;
  var qualifyingTeammates = [];

  otherParams.forEach(function(params) {
    if (params.attack == 1 || params.anomaly == 1) {
      totalQualifyingDamageFocus += params.damageFocus;
      qualifyingTeammates.push({
        params: params,
        weight: params.damageFocus
      });
    }
  });

  if (totalQualifyingDamageFocus <= 0) return;

  // Add to Targets proportionally
  qualifyingTeammates.forEach(function(item) {
    var ratio = item.weight / totalQualifyingDamageFocus;
    var targetAttribute = item.params.attribute;

    // We only add fractional counts for attributes if we want to represent "partial" presence,
    // but typically NumAttribute is an integer count of characters.
    // However, for resonance, maybe it matters?
    // The previous implementation did `NumTarget += 1`.
    // If we split, should we add `ratio` to NumTarget?
    // Let's assume yes, to maintain the logic that "Sunna becomes this attribute".

    if (targetAttribute == "Physical") {
      team.PhysicalAnomalyBuildup += sunnaBuildup * ratio;
      team.PhysicalDamage += sunnaDamage * ratio;
      team.NumPhysical += ratio;
    } else if (targetAttribute == "Ether") {
      team.EtherAnomalyBuildup += sunnaBuildup * ratio;
      team.EtherDamage += sunnaDamage * ratio;
      team.NumEther += ratio;
    } else if (targetAttribute == "Fire") {
      team.FireAnomalyBuildup += sunnaBuildup * ratio;
      team.FireDamage += sunnaDamage * ratio;
      team.NumFire += ratio;
    } else if (targetAttribute == "Ice") {
      team.IceAnomalyBuildup += sunnaBuildup * ratio;
      team.IceDamage += sunnaDamage * ratio;
      team.NumIce += ratio;
    } else if (targetAttribute == "Electric") {
      team.ElectricAnomalyBuildup += sunnaBuildup * ratio;
      team.ElectricDamage += sunnaDamage * ratio;
      team.NumElectric += ratio;
    }
  });
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
 * Compiles a formula string into an executable function.
 * @param {string} expression The formula string (e.g., "1 + damageFocus * 0.5").
 * @return {Function} A function that takes a 'team' object and returns a number.
 */
function compileBuffFunction(expression) {
  if (!expression || expression === "0" || expression.trim() === "") {
    // Return a dummy function that always returns 0 (extremely fast)
    return function() { return 0; };
  }

  // Create a function that takes 'ctx' (context/team) and executes the math
  // We use "Number()" to ensure the result is always a valid number.
  try {
    return new Function("ctx", "with(ctx) { return Number(" + expression + "); }");
  } catch (e) {
    console.error("Failed to compile formula: " + expression, e);
    return function() { return 0; };
  }
}

/**
 * Rounds a number to the nearest 0.25.
 * @param {number} value The value to round.
 * @return {number} The rounded value.
 */
function roundToQuarter(value) {
  return Math.round(value * 4) / 4;
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
 * Helper to safely retrieve or construct a team object.
 * @param {string} char1
 * @param {string} char2
 * @param {string} char3
 * @returns {object|null} The team object or null if invalid.
 */
function getTeamOrCreateSafe(char1, char2, char3) {
  const key = [char1, char2, char3].join("|");
  const teamCharsToTeamObjs = getTeamCharsToTeamObjs();
  if (teamCharsToTeamObjs[key]) {
    return teamCharsToTeamObjs[key];
  }

  // Verify characters exist.
  // Note: char2/char3 might be empty strings if not provided, but addBuffParamsToTeam requires valid params.
  const params = getCharsToBuffParams();
  if (params.has(char1) &&
      params.has(char2) &&
      params.has(char3)) {

     // Create new Class Instance
     const team = new Team(char1, char2, char3);

     // Cache it in the global map.
     teamCharsToTeamObjs[key] = team;
     return team;
  }
  return null;
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

    const team = getTeamOrCreateSafe(character1, character2, character3);

    if (team) {
      var buff = team.calculateBuff(buffExpression);
      // Round to the nearest 0.25.
      buff = roundToQuarter(buff);
      buffs.push([buff]);
    } else {
      buffs.push([0]);
    }
  }
  return buffs;
}

/**
 * Calculates buffs based on the internal team buff formulas of the characters provided.
 * Replaces the need to VLOOKUP formulas in the sheet.
 *
 * @param {Array<Array<string>>} teamRange The range of character names (e.g., F2:H).
 * @param {Array<string>} triggerRange, not used by the formula but invalidates the
 *  cache if the values in the range change.
 * @customfunction
 */
function CALCULATE_TEAM_BUFFS(teamRange, triggerRange) {
  // 1. Initialize data once
  const results = [];

  // 2. Iterate through the rows
  for (let i = 0; i < teamRange.length; i++) {
    const row = teamRange[i];
    const char1 = row[0];

    // Fast exit for empty rows
    if (!char1) {
      break;
    }

    const char2 = row[1];
    const char3 = row[2];

    // 3. Get the team object (calculates combined stats)
    const team = getTeamOrCreateSafe(char1, char2, char3);

    if (!team) {
      break;
    }

    // 4. Retrieve the individual functions we cached in initCharsToBuffParams
    const allParams = getCharsToBuffParams();
    const p1 = allParams.get(char1);
    const p2 = allParams.get(char2);
    const p3 = allParams.get(char3);

    // 5. Execute pre-compiled functions directly.
    let totalBuff = 0;
    if (p1) totalBuff += p1.teamBuffFunction(team);
    if (p2) totalBuff += p2.teamBuffFunction(team);
    if (p3) totalBuff += p3.teamBuffFunction(team);

    results.push([roundToQuarter(totalBuff)]);
  }

  return results;
}

/**
 * Evaluates the javascript buff function based on the columns:
 * Character1 Character2 Character3 char1HasSynergy char2HasSynergy char3HasSynergy char1BuffFunction char2BuffFunction char3BuffFunction
 * * (Synergy, as always, means the given character's Additional Ability is activated.)
 * * @customfunction
 */
function CALCULATE_SYNERGY_BUFFS(data) {
  const results = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const char1 = row[0];

    // Optimization: Break immediately on empty rows. 
    // This is crucial for Sheets inputs like "A2:I" which pass 1000 rows even if 900 are empty.
    if (!char1) break;

    const char2 = row[1];
    const char3 = row[2];

    const teamObj = getTeamOrCreateSafe(char1, char2, char3);

    if (!teamObj) break;

    let totalBuff = 0;

    // Direct index access is faster than destructuring inside a hot loop.
    // Indexes: 3/4/5 are Synergy Booleans, 6/7/8 are Buff Expressions.
    if (row[3]) totalBuff += teamObj.calculateBuff(row[6]);
    if (row[4]) totalBuff += teamObj.calculateBuff(row[7]);
    if (row[5]) totalBuff += teamObj.calculateBuff(row[8]);

    // Round to the nearest 0.25
    results.push([roundToQuarter(totalBuff)]);
  }

  return results;
}
