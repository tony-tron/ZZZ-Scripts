const vm = require('vm');
const fs = require('fs');

const context = vm.createContext({
  console: console,
  Math: Math,
  Map: Map,
  Object: Object,
  String: String,
  Number: Number
});

const buffUtilsCode = fs.readFileSync('09_BuffUtils.js', 'utf8');

vm.runInContext(buffUtilsCode + `
  function main() {
    // Mock the spreadsheet operations
    globalThis.getCharactersDataRange = function() {
      return { getValues: () => [] }; // return empty for now, we will mock getCharsToBuffParams directly
    };
    globalThis.getCharactersColumns = function() { return {}; };
    globalThis.getCharacterNames = function() { return ["A"]; };

    globalThis.getCharsToBuffParams = function() {
      let map = new Map();
      map.set("A", { support: 1, attack: 0, etherDamage: 0, tags: "", fieldTime: 0, stunBuildup: 0, anomalyBuildup: 0,
                     physicalAnomalyBuildup: 0, honedEdgeAnomalyBuildup: 0, etherAnomalyBuildup: 0, fireAnomalyBuildup: 0,
                     iceAnomalyBuildup: 0, frostAnomalyBuildup: 0, electricAnomalyBuildup: 0, windAnomalyBuildup: 0,
                     offFieldDamage: 0, onFieldDamage: 0, damageFocus: 0, physicalDamage: 0, fireDamage: 0, iceDamage: 0,
                     electricDamage: 0, windDamage: 0, sheerDamage: 0, basicAttack: 0, dashAttack: 0, dodgeCounter: 0,
                     assistFollowup: 0, specialAttack: 0, exSpecialAttack: 0, chainAttack: 0, ultimate: 0, anomalyDamage: 0,
                     physicalAnomalyDamage: 0, etherAnomalyDamage: 0, fireAnomalyDamage: 0, iceAnomalyDamage: 0,
                     electricAnomalyDamage: 0, windAnomalyDamage: 0, shieldFocus: 0, healingFocus: 0, quickAssistFocus: 0,
                     chainFocus: 0, chainEnablement: 0, aftershockFocus: 0, exSpecialFocus: 0, aftershockDamage: 0,
                     abloomFocus: 0, abloomDamage: 0, ultimateFocus: 0, ultimateEnablement: 0, hpBenefit: 0, atkBenefit: 0,
                     defBenefit: 0, resShredBenefit: 0, defShredBenefit: 0, impactBenefit: 0, critRateBenefit: 0,
                     critDamageBenefit: 0, energyRegenBenefit: 0, etherVeilFocus: 0,
                     physical: 0, ether: 0, fire: 0, ice: 0, electric: 0, wind: 0, defensiveAssist: 0, evasiveAssist: 0,
                     stun: 0, anomaly: 0, defense: 0, rupture: 0 });
      return map;
    };

    const team = getTeamOrCreateSafe("A", "A", "A");
    try {
        team.calculateBuff("EtherDamage");
        console.log("calculateBuff('EtherDamage') works");
    } catch (e) {
        console.log("Error:", e.message);
    }
  }
  main();
`, context);
