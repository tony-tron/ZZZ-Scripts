const { performance } = require('perf_hooks');

const ITERATIONS = 1_000_000;

// Mock data
const team1 = { strength: 1000 };
const team2 = { strength: 2000 };
const team3 = { strength: 3000 };

const teamTriple = {
  team1Bonus: 100,
  team1ChosenBonus: 10,
  team2Bonus: 200,
  team2ChosenBonus: 20,
  team3Bonus: 300,
  team3ChosenBonus: 30,
  totalStrength: () => 6660,
  minStrength: () => 1000,
};

function testConcat() {
  let result;
  for (let i = 0; i < ITERATIONS; i++) {
    result =
      team1.strength + " + " + teamTriple.team1Bonus + " + " + teamTriple.team1ChosenBonus + "\n+ " +
      team2.strength + " + " + teamTriple.team2Bonus + " + " + teamTriple.team2ChosenBonus + "\n+ " +
      team3.strength + " + " + teamTriple.team3Bonus + " + " + teamTriple.team3ChosenBonus + "\n= " +
      teamTriple.totalStrength() + " (min= " + teamTriple.minStrength() + ")";
  }
  return result;
}

function testTemplate() {
  let result;
  for (let i = 0; i < ITERATIONS; i++) {
    result = `${team1.strength} + ${teamTriple.team1Bonus} + ${teamTriple.team1ChosenBonus}
+ ${team2.strength} + ${teamTriple.team2Bonus} + ${teamTriple.team2ChosenBonus}
+ ${team3.strength} + ${teamTriple.team3Bonus} + ${teamTriple.team3ChosenBonus}
= ${teamTriple.totalStrength()} (min= ${teamTriple.minStrength()})`;
  }
  return result;
}

// Warmup
testConcat();
testTemplate();

const startConcat = performance.now();
testConcat();
const endConcat = performance.now();

const startTemplate = performance.now();
testTemplate();
const endTemplate = performance.now();

const concatTime = endConcat - startConcat;
const templateTime = endTemplate - startTemplate;

console.log(`String concatenation: ${concatTime.toFixed(4)} ms`);
console.log(`Template literals: ${templateTime.toFixed(4)} ms`);
console.log(`Improvement: ${(((concatTime - templateTime) / concatTime) * 100).toFixed(2)}%`);
