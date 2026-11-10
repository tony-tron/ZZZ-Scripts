---
name: zzz-sheet-optimizer
description: Expert system and workflow assistant for the Zenless Zone Zero Team Composition Optimizer spreadsheet and associated Google Apps Script engine.
---

# Zenless Zone Zero Spreadsheet Optimizer Skill

## Purpose & Scope

This skill provides domain knowledge and procedural instructions for maintaining and expanding the **Zenless Zone Zero Team Composition Optimizer** Google Spreadsheet (`1PdatbmxA9f1MXNmv4XCn9BUWFl8ZGVi776Px4VzuwV4`). It covers three primary operational workflows:

1.  **Adding & Evaluating Agents**: Researching released kits or pre-release leaks and translating them into standardized 55-column rows (`A:BC`) on the `Characters` tab.
    *   **Pre-Release & Kit Research Workflow**: When asked to add an unreleased `(TBD)` or newly announced agent without full numbers provided in chat, proactively search pre-release databases (e.g., Zenless Zone Zero Fandom Wiki, Honey Hunter World, Prydwen, Icy Veins, Hakush.in) and official HoYoLAB *Agent Mechanics Intro* posts to gather core stats, passives, energy costs, and recommended Drive Disc 4/5/6 main stats (checking whether guides prioritize PEN Ratio or CRIT DMG due to stat saturation).
    *   **Sheet Formatting & Formula Rule**: Whenever inserting or appending a new agent row on the `Characters` tab, copy the exact cell formatting (including the `Built?` checkbox validation, font family/size, alignment, borders, and number formatting) and row formulas from the preceding character rows—specifically ensuring the **`Damage Focus`** cell (Column `T`) uses the `SUM` formula from previous rows (e.g., `=SUM(U63:AD63)`) rather than a hardcoded number.
2.  **Endgame Buff Formulation**: Translating weekly rotation environmental rules (Shiyu Defense Critical Node / Frontier chambers and Deadly Assault seasonal buffs/bosses) into valid mathematical expressions compatible with the spreadsheet's custom formula engine.
3.  **Automated Rotation Retrieval**: Retrieving active Shiyu Defense and Deadly Assault bosses, elemental profiles, and environment buffs on demand to output standard formulas and lists.

-----

## Part 1: Supported Team Properties & PerChar Properties

### Supported Team Properties (Global Scope)

**Functions & Uptime Helpers:**

  * `AnomalyBuffUptime(uptimeSeconds)`
  * `PerAttributeAnomalyBuffUptime(uptimeSeconds)`
  * `DisorderBuffUptime(uptimeSeconds)`
  * `EXSpecialBuffUptime(uptimeSeconds)`
  * `UltimateBuffUptime(uptimeSeconds)`
  * `StunBuffUptime(uptimeSeconds)`
  * `ChainBuffUptime` (Function)
  * `StunDamageMultiplierOutsideWindow` (Function)
  * `StunDamageMultiplier` (Function)
  * `AbloomBuffUptime` (Function)
  * `ContaminationBuffUptime` (Function)
  * `PerChar('expression')`
  * `MaxChar('expression')`
  * `Buff('attributes')` (e.g. `Buff('ice, wind')`)
  * `Nerf('attributes')` (e.g. `Nerf('physical')`)

**Squad Counts (Always integers 0 to 3; `NumArmorer` is 0 to 1):**

  * Specialties: `NumSupport`, `NumStun`, `NumAttack`, `NumAnomaly`, `NumDefense`, `NumRupture`, `NumArmorer`
  * Attributes: `NumPhysical`, `NumEther`, `NumFire`, `NumIce`, `NumElectric`, `NumWind`, `NumLumiflux`
  * Assists: `NumDefensiveAssist`, `NumEvasiveAssist`

**Team Stats & Aggregations:**

  * General & Stun: `Tags` (String, e.g. `Tags.includes("CritAnomaly")`), `FieldTime`, `StunBuildup`, `StunExtension`, `StunUptime`
  * Anomaly Buildup: `TotalAnomalyBuildup`, `PhysicalAnomalyBuildup`, `HonedEdgeAnomalyBuildup`, `EtherAnomalyBuildup`, `FireAnomalyBuildup`, `IceAnomalyBuildup`, `FrostAnomalyBuildup`, `ElectricAnomalyBuildup`, `WindAnomalyBuildup`, `LumifluxAnomalyBuildup`
  * Damage Focus & Attribute Damage: `OffFieldDamage`, `OnFieldDamage`, `TotalDamageFocus`, `PhysicalDamage`, `EtherDamage`, `FireDamage`, `IceDamage`, `ElectricDamage`, `WindDamage`, `LumifluxDamage`, `SheerDamage`
  * Action Damage: `BasicAttackDamage`, `DashAttackDamage`, `DodgeCounterDamage`, `AssistFollowupDamage`, `SpecialAttackDamage`, `EXSpecialAttackDamage`, `ChainDamage`, `UltimateDamage`
  * Anomaly Damage & Focus: `TotalAnomalyDamage` (NOTE: `AnomalyDmg` is NOT a valid property; always use `TotalAnomalyDamage` at team scope or `anomalyDamage` inside `PerChar`), `PhysicalAnomalyDamage`, `EtherAnomalyDamage`, `FireAnomalyDamage`, `IceAnomalyDamage`, `ElectricAnomalyDamage`, `WindAnomalyDamage`, `LumifluxAnomalyDamage`, `TotalAnomaly`, `HasAttributeAnomaly`, `VortexFocus`, `DisorderFocus`
  * Contamination: `PhysicalContamination`, `EtherContamination`, `FireContamination`, `IceContamination`, `ElectricContamination`, `LumifluxContamination`, `ContaminationDamageBonus`, `ContaminationAnomalyBuildup`
  * Focus & Enablement: `ShieldFocus`, `HealingFocus`, `QuickAssistFocus`, `ForwardAssistFocus`, `BackwardAssistFocus`, `ChainFocus`, `ChainEnablement`, `AftershockFocus`, `EXSpecialFocus`, `EnergyRegenBenefit`, `SharpnessRegenBenefit`, `HasAftershock`, `AftershockDamage`, `AbloomFocus`, `AbloomDamage`, `TeamSheerForce`, `UltimateFocus`, `UltimateEnablement`, `CritRateFocus`, `EtherVeilFocus`, `GashBuildup`
  * Stat Benefits: `HPBenefit`, `AttackBenefit`, `DefenseBenefit`, `ResistanceShredBenefit`, `DefenseShredBenefit`, `DamageBonusBenefit`, `LacerationDamageBonusBenefit`, `SharpDamageBonusBenefit`, `ImpactBenefit`, `CritRateBenefit`, `CritDamageBenefit`

### PerChar Properties (Used strictly inside `PerChar('...')` or `MaxChar('...')`)

  * Identity & Role: `name`, `specialty` (boolean flags: `support`, `stun`, `attack`, `anomaly`, `defense`, `rupture`, `armorer`), `attribute` (boolean flags: `physical`, `ether`, `fire`, `ice`, `electric`, `wind`, `lumiflux`), `defensiveAssist`, `evasiveAssist`, `teamBuff` (Function), `tags`
  * Field & Stun: `fieldTime`, `stunBuildup`, `stunExtension`
  * Anomaly Buildup: `anomalyBuildup`, `physicalAnomalyBuildup`, `honedEdgeAnomalyBuildup`, `etherAnomalyBuildup`, `fireAnomalyBuildup`, `iceAnomalyBuildup`, `frostAnomalyBuildup`, `electricAnomalyBuildup`, `windAnomalyBuildup`, `lumifluxAnomalyBuildup`
  * Damage & Actions: `damageFocus`, `offFieldDamage`, `onFieldDamage`, `physicalDamage`, `etherDamage`, `fireDamage`, `iceDamage`, `electricDamage`, `windDamage`, `lumifluxDamage`, `sheerDamage`, `basicAttack`, `dashAttack`, `dodgeCounter`, `assistFollowup`, `specialAttack`, `exSpecialAttack`, `chainAttack`, `ultimate`
  * Anomaly Damage: `anomalyDamage`, `physicalAnomalyDamage`, `etherAnomalyDamage`, `fireAnomalyDamage`, `iceAnomalyDamage`, `electricAnomalyDamage`, `windAnomalyDamage`, `lumifluxAnomalyDamage`
  * Focus & Enablement: `shieldFocus`, `healingFocus`, `quickAssistFocus`, `chainFocus`, `chainEnablement`, `aftershockFocus`, `aftershockDamage`, `exSpecialFocus`, `exSpecialBuffUptime`, `energyRegenBenefit`, `sharpnessRegenBenefit`, `ultimateFocus`, `ultimateBuffUptime`, `ultimateEnablement`, `critRateFocus`, `etherVeilFocus`, `abloomFocus`, `abloomDamage`, `teamSheerForce`, `gashBuildup`
  * Benefits: `hpBenefit`, `atkBenefit`, `defBenefit`, `resShredBenefit`, `defShredBenefit`, `damageBonusBenefit`, `lacerationDamageBonusBenefit`, `sharpDamageBonusBenefit`, `impactBenefit`, `critRateBenefit`, `critDamageBenefit`

-----

## Part 2: Core Estimation Baselines & Calibration Rules (55-Column `Characters` Tab, `A:BC`)

  * **Base Strength**: **MUST ALWAYS be a multiple of 0.25** (e.g., `2.50`, `3.00`, `4.00`, `4.25`).
  * **Synergy Bonus (Additional Ability)**: **Almost never blank**. Formulate the exact buff expression or flat strength value granted when Additional Ability requirements are met (e.g., `CritDamageBenefit * 0.15` or `0.7` for a personal flat ~840–850 ATK self-buff).
  * **Extra Character Synergy & Extra Synergy Bonus (`Columns L–M`)**:
      * Default to blank in `L` and `1,1` in `M`.
      * When a Core Passive or kit mechanic grants a **fixed one-off bonus** for having teammates of a specific Attribute, Specialty, or stat threshold (without needing dynamic team-variable math), encode the filter in `Extra Character Synergy` (e.g., `Attribute='Wind'`, `Specialty='Stun'`) and set `Extra Synergy Bonus` to `<1_match_bonus>,<2_matches_bonus>` (e.g., `0.6,0` when only one matching teammate is needed, or `0,1` when two are required).
  * **Team Synergy (Core Passive & Global Buffs)**: Formulate dynamic team-wide expressions provided to or scaling off the squad (e.g., `Math.sqrt(StunUptime)*0.5 + UltimateEnablement*0.5 + ChainEnablement*0.25`).
  * **Field Time**: **ALWAYS 2.0 or less** (`0` off-field, `1.0` quickswap/hybrid, `2.0` hypercarry).
  * **Stun Buildup & Stun Extension**: If there is no daze amplification, leave blank / `0`.
  * **Damage Focus & Action Allocations (`Columns T–AD`)**:
      * **`Damage Focus` (`Column T`) MUST use the `SUM` formula copied from previous rows** (`=SUM(U<row>:AD<row>)`, e.g., `=SUM(U63:AD63)`) rather than a hardcoded number.
      * Populate the sub-allocations across `Basic Attack` through `Other?` (`Columns U–AD`) so their sum equals `0`, `1`, `1.5`, or `2`. Leave negligible action weights blank.
  * **Anomaly Buildup**: `1.0` for non-anomaly; `1.5`, `2.0`, or `2.5` for high/anomaly units.
  * **Ultimate Focus**: Default is `1.0` (step up to `1.25`–`2.0` if Ultimate refunds major core resources or dominates rotation burst).
  * **Energy Regen Benefit**: When EX Special energy cost is $X$, calibrate as approximately $20 / X$.
  * **Stat Benefits Calibration (Scaled by `Damage Focus = D`, rounded to nearest `0.25`)**:
      * **`Dmg Bonus Benefit`**: Default equals **`D`** (`2.0` for main carries, `0` for zero-damage supports/stunners). Step down slightly only under extreme DMG% saturation.
      * **`Res Shred Benefit` & `DEF Shred Benefit`**:
          * **Baseline Rule**: Default both to **`0.5 * D`** (`1.0` for `D = 2.0` carries, `0.5` for `D = 1.0`).
          * **Rupture Exception**: Set `DEF Shred Benefit = 0` (Sheer DMG ignores DEF) while keeping `Res Shred Benefit = 0.5 * D`.
          * **Saturation Exception**: Increase up to `1.5` or `2.0` only as an exception when theorycrafting guides highlight heavy saturation in DMG Bonus / ATK / CRIT (e.g., favoring PEN Ratio on Disc 5 and prioritizing DEF/RES shred).
      * **Base Scaling Stat (`ATK` / `HP` / `DEF Benefit`)**:
          * For `D = 2.0` carries: Attack = `ATK Benefit: 1.0`; Anomaly = `ATK Benefit: 0.5` (`1.0` for Crit/ATK-heavy Anomaly); Rupture = `HP Benefit: 1.0`, `ATK Benefit: 0.5`; Armorer = `DEF Benefit: 1.5`. Scale proportionally for lower `D` (usually `0`, `0.5`, or `1.0`).
      * **CRIT Columns & Self-Buff Adjustments (`Crit Rate Benefit`, `Crit DMG Benefit`, `Crit Rate Focus`)**:
          * **Baseline (`D = 2.0` Crit Carry)**: `Crit Rate Benefit = 1.0`, `Crit DMG Benefit = 1.0`, `Crit Rate Focus = 0.75` (`1.0` if routinely operating at ~90–100% effective CRIT Rate).
          * **Self-Buff Bias vs. Scaling Synergy**: Bias slightly toward unbuffed stats when one bucket is heavily saturated, **with caution**—kits often self-buff the exact stat they scale hardest with. Always check passives first:
              * *In-Combat Conversions / Overflow (Raise to `1.5`)*: Increase benefit when an agent converts an **in-combat** stat into team/personal multipliers or scales past 100% CRIT Rate (e.g., Soldier 0 Anby's CRIT DMG -> Aftershock CRIT DMG; Claret's >100% CRIT Rate -> multi-CRIT Laceration). *Initial* (pre-combat stat screen) conversions do not scale with in-combat buffs.
              * *100% CRIT Rate Cap Proximity (Lower `Crit Rate Benefit` to `0.5`, optionally raise `Crit DMG Benefit` to `1.25`–`1.5`)*: Apply when massive free CRIT Rate makes external CRIT Rate buffs likely to overcap 100% without overflow scaling.

-----

## Part 3: Quantitative Conversion & Formula Syntax Rules

1.  **Boss Modifiers**: `Buff()` and `Nerf()` accept a single comma-separated lowercase string argument:
      * `Buff('ice, wind') + Nerf('physical')`
2.  **Specialty Scoping**: `Specialty` is NOT a team attribute. To scope an effect to a specialty (e.g. Stun or Attack), use `PerChar('stun * (...)')` or `PerChar('specialty == "Stun" ? ... : 0')`.
3.  **Tag-Based Triggers**:
      * For mechanics tied to specific kit tags (such as triggering critical hits on Attribute Anomaly damage), check the team `Tags` string via `Tags.includes("CritAnomaly")` (e.g., `Tags.includes("CritAnomaly") * (TotalAnomalyDamage * 0.20 + FireAnomalyDamage * 0.10)`).
4.  **Squad Count Array Indexing**:
      * `NumAnomaly` (and other `Num*` variables) is always `0`, `1`, `2`, or `3`. Do NOT append `|| 0` to 4-element array lookups:
      * `[0, 0, 0.05, 0.125][NumAnomaly] * TotalAnomalyDamage`
5.  **Stun Recovery Speed & Stunned Enemy Damage**:
      * "Upon hitting a Stunned enemy, their Stun recovery Speed reduces by 10% for 15s": encode as `StunBuffUptime(15) * StunUptime * 0.1`.
      * Bonus damage against Stunned enemies (e.g. +25% Anomaly DMG against stunned enemies): multiply by `StunUptime * 0.1` (or `StunUptime * 0.11` when paired with a 10% stun recovery speed reduction in the same buff, e.g., `TotalAnomalyDamage * 0.25 * StunUptime * 0.11`).
6.  **Uptime Helpers**: Wrap combined additive uptimes in `Math.min(1, ...)`:
      * `Math.min(1, DisorderBuffUptime(15) + AbloomBuffUptime(15)) * (...)`
7.  **Decibel Gains & Ultimate Value**:
      * External Decibel restoration (such as environment buffs granting 500 Decibels on defeating an Elite or completing an objective) scales directly with how much value the team derives from fielding Ultimates. Encode this using `UltimateFocus` (e.g., `UltimateFocus * (500 / 3000)` or `UltimateFocus * 0.15`), NOT `UltimateEnablement` (which strictly quantifies an agent's kit-inherent ability to feed Decibels or provide extra ultimate opportunities to teammates).
