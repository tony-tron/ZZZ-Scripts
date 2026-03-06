const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

// Load script
const code = fs.readFileSync('09_BuffUtils.js', 'utf8');

// Mock global objects and functions needed by the script
const context = {
  console: console,
  _charsToBuffParams: null,
  getCharactersDataRange: () => ({
    getValues: () => [
      ['Headers'],
      ['Sunna', 'Attack', 'Physical', 'Defensive', 10, 5, 100, 20, 5, ...Array(25).fill(0)], // Fake char data
      ['Jane', 'Anomaly', 'Electric', 'Evasive', 15, 6, 80, 50, 10, ...Array(25).fill(0)],
      ['Support', 'Support', 'Ice', 'Evasive', 5, 0, 10, 5, 0, ...Array(25).fill(0)]
    ]
  }),
  getCharactersColumns: () => ({
    specialty: 1, attribute: 2, assistType: 3, anomalyBuildup: 4, fieldTime: 5, damageFocus: 6, anomalyDamage: 7, aftershockFocus: 8, teamBuffFormula: 9, tags: 10,
    stunBuildup: 11, basicAttack: 12, dashAttack: 13, dodgeCounter: 14, assistFollowup: 15, specialAttack: 16, exSpecialAttack: 17, chainAttack: 18, ultimate: 19,
    shieldFocus: 20, healingFocus: 21, quickAssistFocus: 22, chainFocus: 23, chainEnablement: 24, exSpecialFocus: 25, ultimateFocus: 26, ultimateEnablement: 27,
    hpBenefit: 28, atkBenefit: 29, defBenefit: 30, resShredBenefit: 31, defShredBenefit: 32, impactBenefit: 33, critRateBenefit: 34, critDamageBenefit: 35, energyRegenBenefit: 36, etherVeilFocus: 37
  }),
  getCharacterNames: () => ['Sunna', 'Jane', 'Support'],
  _updateTeamForYuzuha: () => {}, // Mocked out
  getTeamCharsToTeamObjs: () => ({}),
  Map: Map,
  Math: Math,
  Number: Number,
  String: String,
  Array: Array
};

vm.createContext(context);
vm.runInContext(code, context);

// Test setup
vm.runInContext(`
  // Manual override of getCharsToBuffParams for testing
  const charsMap = new Map();
  charsMap.set('Sunna', {
    name: 'Sunna', attribute: 'Physical', attack: 1, anomaly: 0, damageFocus: 100, anomalyBuildup: 10, anomalyDamage: 50,
    physicalDamage: 100, physicalAnomalyBuildup: 10, physicalAnomalyDamage: 50,
    electricDamage: 0, electricAnomalyBuildup: 0, electricAnomalyDamage: 0,
    iceDamage: 0, iceAnomalyBuildup: 0, iceAnomalyDamage: 0
  });
  charsMap.set('Jane', {
    name: 'Jane', attribute: 'Electric', attack: 0, anomaly: 1, damageFocus: 80, anomalyBuildup: 20, anomalyDamage: 40,
    physicalDamage: 0, physicalAnomalyBuildup: 0, physicalAnomalyDamage: 0,
    electricDamage: 80, electricAnomalyBuildup: 20, electricAnomalyDamage: 40,
    iceDamage: 0, iceAnomalyBuildup: 0, iceAnomalyDamage: 0
  });
  charsMap.set('Support', {
    name: 'Support', attribute: 'Ice', attack: 0, anomaly: 0, damageFocus: 10, anomalyBuildup: 5, anomalyDamage: 5,
    physicalDamage: 0, physicalAnomalyBuildup: 0, physicalAnomalyDamage: 0,
    electricDamage: 0, electricAnomalyBuildup: 0, electricAnomalyDamage: 0,
    iceDamage: 10, iceAnomalyBuildup: 5, iceAnomalyDamage: 5
  });
  getCharsToBuffParams = () => charsMap;
`, context);

// Test Case 1: Sunna with a qualifying teammate (Jane - Anomaly)
vm.runInContext(`
  const team1 = new Team('Sunna', 'Jane', 'Support');
  // Check that Sunna's stats are removed from Physical
  if (team1.PhysicalAnomalyBuildup !== 0) throw new Error("PhysicalAnomalyBuildup not removed: " + team1.PhysicalAnomalyBuildup);
  if (team1.PhysicalDamage !== 0) throw new Error("PhysicalDamage not removed: " + team1.PhysicalDamage);
  if (team1.PhysicalAnomalyDamage !== 0) throw new Error("PhysicalAnomalyDamage not removed: " + team1.PhysicalAnomalyDamage);

  // Check that Sunna's stats are added to Electric (Jane's attribute)
  // Jane's damageFocus = 80. Total qualifying = 80. Ratio = 80/80 = 1.
  if (team1.ElectricAnomalyBuildup !== 30) throw new Error("ElectricAnomalyBuildup not updated: " + team1.ElectricAnomalyBuildup); // Jane's 20 + Sunna's 10
  if (team1.ElectricDamage !== 180) throw new Error("ElectricDamage not updated: " + team1.ElectricDamage); // Jane's 80 + Sunna's 100
  if (team1.ElectricAnomalyDamage !== 90) throw new Error("ElectricAnomalyDamage not updated: " + team1.ElectricAnomalyDamage); // Jane's 40 + Sunna's 50
`, context);

// Test Case 2: Sunna without qualifying teammates (Only Support)
vm.runInContext(`
  charsMap.set('Support2', {
    name: 'Support2', attribute: 'Ice', attack: 0, anomaly: 0, damageFocus: 10, anomalyBuildup: 5, anomalyDamage: 5,
    physicalDamage: 0, physicalAnomalyBuildup: 0, physicalAnomalyDamage: 0,
    electricDamage: 0, electricAnomalyBuildup: 0, electricAnomalyDamage: 0,
    iceDamage: 10, iceAnomalyBuildup: 5, iceAnomalyDamage: 5
  });

  const team2 = new Team('Sunna', 'Support', 'Support2');
  // Total qualifying damage focus is 0. Sunna's stats should NOT be removed.
  if (team2.PhysicalAnomalyBuildup !== 10) throw new Error("PhysicalAnomalyBuildup improperly removed: " + team2.PhysicalAnomalyBuildup);
  if (team2.PhysicalDamage !== 100) throw new Error("PhysicalDamage improperly removed: " + team2.PhysicalDamage);
  if (team2.PhysicalAnomalyDamage !== 50) throw new Error("PhysicalAnomalyDamage improperly removed: " + team2.PhysicalAnomalyDamage);
`, context);

console.log('All tests passed!');
