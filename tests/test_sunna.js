const fs = require('fs');
const path = require('path');

// Read BuffUtils.js
const buffUtilsPath = path.join(__dirname, '../BuffUtils.js');
let buffUtilsContent = fs.readFileSync(buffUtilsPath, 'utf8');

// Bypass initCharsToBuffParams execution
buffUtilsContent = buffUtilsContent.replace(
    'var charsToBuffParams = initCharsToBuffParams();',
    'var charsToBuffParams = new Map();'
);

// Mock global variables
global.charsToBuffParams = new Map();

// Helper to create dummy params
function createParams(name, attribute, damageFocus, anomalyDamage, anomalyBuildup) {
    return {
        name: name,
        attribute: attribute,
        damageFocus: damageFocus || 0,
        anomalyDamage: anomalyDamage || 0,
        anomalyBuildup: anomalyBuildup || 0,

        // Flags needed for addBuffParamsToTeam
        support: 0, stun: 0, attack: 1, anomaly: 0, defense: 0, rupture: 0,
        physical: attribute === "Physical" ? 1 : 0,
        ether: attribute === "Ether" ? 1 : 0,
        fire: attribute === "Fire" ? 1 : 0,
        ice: attribute === "Ice" ? 1 : 0,
        electric: attribute === "Electric" ? 1 : 0,

        defensiveAssist: 0, evasiveAssist: 1,
        tags: "Test",
        fieldTime: 0,
        stunBuildup: 0,
        sheerDamage: 0,

        // Specific anomaly buildups (simplified logic from initCharsToBuffParams)
        physicalAnomalyBuildup: attribute === "Physical" ? anomalyBuildup : 0,
        etherAnomalyBuildup: attribute === "Ether" ? anomalyBuildup : 0,
        fireAnomalyBuildup: attribute === "Fire" ? anomalyBuildup : 0,
        iceAnomalyBuildup: attribute === "Ice" ? anomalyBuildup : 0,
        electricAnomalyBuildup: attribute === "Electric" ? anomalyBuildup : 0,
        honedEdgeAnomalyBuildup: 0, frostAnomalyBuildup: 0,

        offFieldDamage: damageFocus,
        onFieldDamage: 0,

        // Damage breakdown
        physicalDamage: attribute === "Physical" ? damageFocus : 0,
        etherDamage: attribute === "Ether" ? damageFocus : 0,
        fireDamage: attribute === "Fire" ? damageFocus : 0,
        iceDamage: attribute === "Ice" ? damageFocus : 0,
        electricDamage: attribute === "Electric" ? damageFocus : 0,

        basicAttack: 0, dashAttack: 0, dodgeCounter: 0, assistFollowup: 0,
        specialAttack: 0, exSpecialAttack: 0, chainAttack: 0, ultimate: 0,

        // Anomaly damage
        physicalAnomalyDamage: attribute === "Physical" ? anomalyDamage : 0,
        etherAnomalyDamage: attribute === "Ether" ? anomalyDamage : 0,
        fireAnomalyDamage: attribute === "Fire" ? anomalyDamage : 0,
        iceAnomalyDamage: attribute === "Ice" ? anomalyDamage : 0,
        electricAnomalyDamage: attribute === "Electric" ? anomalyDamage : 0,

        shieldFocus: 0, healingFocus: 0, quickAssistFocus: 0,
        chainFocus: 0, chainEnablement: 0, aftershockFocus: 0, aftershockDamage: 0,
        exSpecialFocus: 0, ultimateFocus: 0, ultimateEnablement: 0,
        hpBenefit: 0, atkBenefit: 0, defBenefit: 0, resShredBenefit: 0, defShredBenefit: 0,
        impactBenefit: 0, critRateBenefit: 0, critDamageBenefit: 0, energyRegenBenefit: 0
    };
}

// Evaluate BuffUtils.js
// We wrap in a function to avoid const redeclaration issues if run multiple times,
// but here we run once. However, `const formulaCache` at top level might be an issue if we just eval.
// Since we are running in node, top-level variables in eval become local to the eval scope unless attached to global.
// BuffUtils.js defines functions which we need.
eval(buffUtilsContent);

// --- Test Case 1: Sunna adapts to Stronger Teammate ---

console.log("--- Test Case 1 ---");
const sunna = createParams("Sunna", "Physical", 100, 50, 20); // Base Physical
const ally1 = createParams("Ally1", "Fire", 200, 100, 30); // Stronger
const ally2 = createParams("Ally2", "Ice", 50, 20, 10);   // Weaker

charsToBuffParams.set("Sunna", sunna);
charsToBuffParams.set("Ally1", ally1);
charsToBuffParams.set("Ally2", ally2);

const team = {
    characters: ["Sunna", "Ally1", "Ally2"],
};

try {
    addBuffParamsToTeam(team);

    console.log("Sunna Attribute: Physical");
    console.log("Ally1 Attribute: Fire (Target)");
    console.log("Ally2 Attribute: Ice");

    console.log("Team Physical Anomaly Buildup:", team.PhysicalAnomalyBuildup);
    console.log("Team Fire Anomaly Buildup:", team.FireAnomalyBuildup);
    console.log("Team NumPhysical:", team.NumPhysical);
    console.log("Team NumFire:", team.NumFire);

    // After change expectation:
    // Sunna (20) moves to Fire.
    // Physical: 0 (was 20)
    // Fire: 30 + 20 = 50.
    // NumPhysical: 0.
    // NumFire: 2.

    if (team.PhysicalAnomalyBuildup === 0 && team.FireAnomalyBuildup === 50 && team.NumPhysical === 0 && team.NumFire === 2) {
        console.log("SUCCESS: Sunna adapted to Fire.");
    } else if (team.PhysicalAnomalyBuildup === 20 && team.FireAnomalyBuildup === 30) {
        console.log("BASELINE: Sunna stayed Physical (Not Implemented Yet).");
    } else {
        console.log("FAIL: Unexpected values.");
    }

} catch (e) {
    console.error("Error executing addBuffParamsToTeam:", e);
}
