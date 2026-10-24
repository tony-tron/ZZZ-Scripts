/** @OnlyCurrentDoc */

var _sortedTeamsCache = null;

function clearSortedTeamsCache() {
  _sortedTeamsCache = null;
}

function getAllTeams(minTeamStrength) {
  if (!_sortedTeamsCache) {
    const rangeStr = typeof sortedTeamsRangeString !== "undefined" ? sortedTeamsRangeString : "A2:D";
    const data = getSortedTeamsSheet().getRange(rangeStr).getValues();
    const allTeams = [];
    for (let r = 0; r < data.length; r++) {
      const char1 = data[r][0];
      if (!char1) break;
      const strength = Number(data[r][3]);
      if (isNaN(strength)) continue;
      const team = new Team(char1, data[r][1], data[r][2]);
      team.strength = strength;
      allTeams.push(team);
    }
    _sortedTeamsCache = allTeams;
  }

  if (typeof minTeamStrength === "number" && !isNaN(minTeamStrength) && minTeamStrength > 0) {
    return _sortedTeamsCache.filter(team => team.strength >= minTeamStrength);
  }
  return _sortedTeamsCache;
}

/**
 * Core Branch-and-Bound optimizer for finding the top k-team distinct combinations
 * across all available teams.
 *
 * Precomputes all team buff bonuses and selectable buff bonuses once per team,
 * uses 64-bit BigInt character bitmasks for O(1) overlap checks, and optimizes
 * room permutations directly for the target sheet sort metric.
 *
 * @param {Array<Team>} teams - The list of all available teams.
 * @param {Array<Array<string>>} buffExpressionsList - Array of k buff expression arrays (one per slot).
 * @param {Array<object>} [buffOptions] - Optional list of selectable buff options ({ name, expression }).
 * @param {number} [maxOptions=50] - Maximum number of top combinations to return.
 * @returns {Array<object>} Sorted list of top distinct team combination objects.
 */
function computeBestDistinctTeamGroups(
  teams,
  buffExpressionsList,
  buffOptions,
  maxOptions = 50
) {
  const k = buffExpressionsList ? buffExpressionsList.length : 0;
  const n = teams ? teams.length : 0;
  if (k === 0 || n < k) return [];

  const limit = (typeof maxOptions === "number" && maxOptions > 0) ? maxOptions : 50;
  const hasBuffOptions = Array.isArray(buffOptions) && buffOptions.length > 0;

  // 1. Ensure 64-bit character bitmasks (as two 32-bit ints) exist on all teams
  const charMasksLo = new Int32Array(n);
  const charMasksHi = new Int32Array(n);
  let fallbackCharMap = null;
  for (let i = 0; i < n; i++) {
    const t = teams[i];
    if (typeof t.charMaskLo === "number") {
      charMasksLo[i] = t.charMaskLo;
      charMasksHi[i] = t.charMaskHi;
    } else {
      if (!fallbackCharMap) fallbackCharMap = new Map();
      let maskLo = 0;
      let maskHi = 0;
      for (let c = 0; c < t.characters.length; c++) {
        const name = t.characters[c];
        let cid = fallbackCharMap.get(name);
        if (cid === undefined) {
          cid = fallbackCharMap.size;
          fallbackCharMap.set(name, cid);
        }
        if (cid < 32) {
          maskLo |= (1 << cid);
        } else {
          maskHi |= (1 << (cid - 32));
        }
      }
      charMasksLo[i] = maskLo;
      charMasksHi[i] = maskHi;
    }
  }

  // 2. Precompute best chosen buff per team (once per team)
  const chosenBonus = new Float64Array(n);
  const chosenName = new Array(n);
  if (hasBuffOptions) {
    const numOpts = buffOptions.length;
    const optExprs = new Array(numOpts);
    const optNames = new Array(numOpts);
    for (let b = 0; b < numOpts; b++) {
      optExprs[b] = buffOptions[b].expression;
      optNames[b] = buffOptions[b].name;
    }
    for (let i = 0; i < n; i++) {
      const team = teams[i];
      let bestVal = 0;
      let bestName = optNames[0];
      for (let b = 0; b < numOpts; b++) {
        const val = team.calculateBuff(optExprs[b]);
        if (val > bestVal) {
          bestVal = val;
          bestName = optNames[b];
        }
      }
      chosenBonus[i] = Math.round(bestVal * 1000) / 1000;
      chosenName[i] = bestName;
    }
  } else {
    chosenName.fill("");
  }

  // 3. Precompute slot bonuses and effective strengths per slot, deduplicating identical slots
  const slotBonus = new Array(k);
  const slotEff = new Array(k);
  const slotSigs = buffExpressionsList.map(exprs => JSON.stringify((exprs || []).map(String)));

  for (let s = 0; s < k; s++) {
    let dupOf = -1;
    for (let prev = 0; prev < s; prev++) {
      if (slotSigs[prev] === slotSigs[s]) {
        dupOf = prev;
        break;
      }
    }
    if (dupOf !== -1) {
      slotBonus[s] = slotBonus[dupOf];
      slotEff[s] = slotEff[dupOf];
      continue;
    }

    const bonuses = new Float64Array(n);
    const effs = new Float64Array(n);
    const exprs = buffExpressionsList[s] || [];
    const numExprs = exprs.length;

    for (let i = 0; i < n; i++) {
      const team = teams[i];
      let added = 0;
      for (let e = 0; e < numExprs; e++) {
        added += team.calculateBuff(exprs[e]);
      }
      const roundedBonus = Math.round(added * 1000) / 1000;
      bonuses[i] = roundedBonus;
      effs[i] = Math.round((team.strength + roundedBonus + chosenBonus[i]) * 1000) / 1000;
    }
    slotBonus[s] = bonuses;
    slotEff[s] = effs;
  }

  // Fast path for k = 1
  if (k === 1) {
    const order = Int32Array.from({ length: n }, (_, i) => i);
    const eff0 = slotEff[0];
    order.sort((a, b) => eff0[b] - eff0[a] || a - b);
    const count = Math.min(limit, n);
    const results = new Array(count);
    for (let idx = 0; idx < count; idx++) {
      const i = order[idx];
      const tot = eff0[i];
      results[idx] = {
        team1: teams[i],
        team1Bonus: slotBonus[0][i],
        team1ChosenBonusName: chosenName[i],
        team1ChosenBonus: chosenBonus[i],
        totalStrength: function() { return tot; },
        minStrength: function() { return tot; },
      };
    }
    return results;
  }

  // 4. Sort candidate team indices per slot in descending order of effective strength
  const slotOrder = new Array(k);
  const slotRank = new Array(k);
  const maxSlotEff = new Float64Array(k);

  for (let s = 0; s < k; s++) {
    let dupOf = -1;
    for (let prev = 0; prev < s; prev++) {
      if (slotSigs[prev] === slotSigs[s]) {
        dupOf = prev;
        break;
      }
    }
    if (dupOf !== -1) {
      slotOrder[s] = slotOrder[dupOf];
      slotRank[s] = slotRank[dupOf];
      maxSlotEff[s] = maxSlotEff[dupOf];
      continue;
    }

    const effs = slotEff[s];
    const order = Int32Array.from({ length: n }, (_, i) => i);
    order.sort((a, b) => effs[b] - effs[a] || a - b);
    const rank = new Int32Array(n);
    for (let pos = 0; pos < n; pos++) {
      rank[order[pos]] = pos;
    }
    slotOrder[s] = order;
    slotRank[s] = rank;
    maxSlotEff[s] = effs[order[0]];
  }

  // 5. Order slots for search: group identical slots together and sort groups by maxSlotEff descending
  const searchSlots = Array.from({ length: k }, (_, s) => s);
  searchSlots.sort((a, b) => {
    if (slotSigs[a] === slotSigs[b]) return a - b;
    return maxSlotEff[b] - maxSlotEff[a] || a - b;
  });

  const sameAsPrevSlot = new Uint8Array(k);
  for (let d = 1; d < k; d++) {
    if (slotSigs[searchSlots[d]] === slotSigs[searchSlots[d - 1]]) {
      sameAsPrevSlot[d] = 1;
    }
  }

  // Precompute upper bounds for remaining slots at each depth d
  const remMaxSum = new Float64Array(k);
  const remMaxMin = new Float64Array(k);
  for (let d = 0; d < k; d++) {
    let sum = 0;
    let minVal = Infinity;
    for (let r = d + 1; r < k; r++) {
      const m = maxSlotEff[searchSlots[r]];
      sum += m;
      if (m < minVal) minVal = m;
    }
    remMaxSum[d] = sum;
    remMaxMin[d] = minVal;
  }

  // Precompute all k! permutations for leaf assignment evaluation
  const perms = getAllPermutations(Array.from({ length: k }, (_, i) => i));
  const numPerms = perms.length;

  // Comparison helper: 50% minStrength + 50% averageStrength, tie-breaking by totalStrength then minStrength
  function compareMetrics(scoreA, minA, totA, scoreB, minB, totB) {
    if (Math.abs(scoreA - scoreB) > 1e-9) return scoreB - scoreA;
    if (totA !== totB) return totB - totA;
    return minB - minA;
  }

  // Top-M collection (sorted best-to-worst, capped at `limit`)
  const topResults = [];
  let thresholdScore = -Infinity;
  let thresholdMin = -Infinity;
  let thresholdTot = -Infinity;

  function updateThreshold() {
    if (topResults.length >= limit) {
      const worst = topResults[limit - 1];
      thresholdScore = worst.score;
      thresholdMin = worst.minVal;
      thresholdTot = worst.totVal;
    }
  }

  const visitedSets = new Set();
  const pickedTeams = new Int32Array(k);
  const sortedPicked = new Int32Array(k);

  function evaluateLeaf() {
    // Canonicalize the unordered set of k team indices to avoid duplicate combinations
    for (let i = 0; i < k; i++) sortedPicked[i] = pickedTeams[i];
    sortedPicked.sort();

    let key = "";
    for (let i = 0; i < k; i++) {
      key += (i === 0 ? "" : ",") + sortedPicked[i];
    }
    if (visitedSets.has(key)) return;
    visitedSets.add(key);

    // Find optimal assignment of the k teams across the k slots
    let bestScore = -Infinity;
    let bestMin = -Infinity;
    let bestTot = -Infinity;
    let bestPerm = perms[0];

    for (let p = 0; p < numPerms; p++) {
      const perm = perms[p];
      let tot = 0;
      let minV = Infinity;
      for (let s = 0; s < k; s++) {
        const tIdx = sortedPicked[perm[s]];
        const eff = slotEff[s][tIdx];
        tot += eff;
        if (eff < minV) minV = eff;
      }
      const rTot = Math.round(tot * 1000) / 1000;
      const rMin = Math.round(minV * 1000) / 1000;
      const score = 0.5 * rMin + 0.5 * (rTot / k);

      if (compareMetrics(score, rMin, rTot, bestScore, bestMin, bestTot) < 0) {
        bestScore = score;
        bestMin = rMin;
        bestTot = rTot;
        bestPerm = perm;
      }
    }

    if (
      topResults.length >= limit &&
      compareMetrics(bestScore, bestMin, bestTot, thresholdScore, thresholdMin, thresholdTot) >= 0
    ) {
      return;
    }

    const assignedTeams = new Int32Array(k);
    for (let s = 0; s < k; s++) {
      assignedTeams[s] = sortedPicked[bestPerm[s]];
    }

    const entry = {
      assignedTeams,
      score: bestScore,
      minVal: bestMin,
      totVal: bestTot,
    };

    // Insert into topResults
    let insertIdx = topResults.length;
    while (
      insertIdx > 0 &&
      compareMetrics(
        entry.score,
        entry.minVal,
        entry.totVal,
        topResults[insertIdx - 1].score,
        topResults[insertIdx - 1].minVal,
        topResults[insertIdx - 1].totVal
      ) < 0
    ) {
      insertIdx--;
    }
    topResults.splice(insertIdx, 0, entry);
    if (topResults.length > limit) {
      topResults.pop();
    }
    updateThreshold();
  }

  function search(depth, usedLo, usedHi, curSum, curMin, prevPosInSameSlotGroup) {
    const slot = searchSlots[depth];
    const order = slotOrder[slot];
    const effs = slotEff[slot];
    const startPos = sameAsPrevSlot[depth] ? prevPosInSameSlotGroup + 1 : 0;
    const remSum = remMaxSum[depth];
    const remMin = remMaxMin[depth];
    const isLast = depth === k - 1;

    for (let pos = startPos; pos < n; pos++) {
      const candEff = effs[order[pos]];

      // Branch-and-Bound pruning: since order is sorted descending by candEff,
      // both upperTotal and upperMin decrease monotonically as pos increases.
      if (topResults.length >= limit) {
        const upperTot = curSum + candEff + remSum;
        const upperMin = Math.min(curMin, candEff, remMin);
        const upperScore = 0.5 * upperMin + 0.5 * (upperTot / k);
        if (upperScore < thresholdScore - 1e-9) break;
        if (
          Math.abs(upperScore - thresholdScore) <= 1e-9 &&
          upperTot < thresholdTot - 1e-9
        ) {
          break;
        }
      }

      const tIdx = order[pos];
      const tLo = charMasksLo[tIdx];
      const tHi = charMasksHi[tIdx];
      if ((usedLo & tLo) !== 0 || (usedHi & tHi) !== 0) continue;

      pickedTeams[depth] = tIdx;
      const nextSum = curSum + candEff;
      const nextMin = candEff < curMin ? candEff : curMin;

      if (isLast) {
        evaluateLeaf();
      } else {
        search(depth + 1, usedLo | tLo, usedHi | tHi, nextSum, nextMin, pos);
      }
    }
  }

  search(0, 0, 0, 0, Infinity, -1);

  // 6. Materialize output objects matching the exact schema expected by all callers
  return topResults.map(entry => {
    const obj = {};
    for (let s = 0; s < k; s++) {
      const tIdx = entry.assignedTeams[s];
      const num = s + 1;
      obj[`team${num}`] = teams[tIdx];
      obj[`team${num}Bonus`] = slotBonus[s][tIdx];
      obj[`team${num}ChosenBonusName`] = chosenName[tIdx];
      obj[`team${num}ChosenBonus`] = chosenBonus[tIdx];
    }
    obj.totalStrength = function() {
      let sum = 0;
      for (let s = 1; s <= k; s++) {
        sum += this[`team${s}`].strength + (this[`team${s}Bonus`] || 0) + (this[`team${s}ChosenBonus`] || 0);
      }
      return Math.round(sum * 1000) / 1000;
    };
    obj.minStrength = function() {
      let minV = Infinity;
      for (let s = 1; s <= k; s++) {
        const eff = this[`team${s}`].strength + (this[`team${s}Bonus`] || 0) + (this[`team${s}ChosenBonus`] || 0);
        if (eff < minV) minV = eff;
      }
      return Math.round(minV * 1000) / 1000;
    };
    return obj;
  });
}

/**
 * Computes 1 distinct team, for completeness (all teams are distinct so why would you want this?).
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<string>} team1BuffExpressions - Buffs for slot 1.
 * @param {Array<object>} [buffOptions] - Selectable buff options.
 * @param {number} [maxOptions=50] - Max options to return.
 * @returns {Array<object>} A list of optimized 1-team objects.
 */
function computeBestDistinctTeamSingles(teams, team1BuffExpressions, buffOptions, maxOptions = 50) {
  return computeBestDistinctTeamGroups(teams, [team1BuffExpressions], buffOptions, maxOptions);
}

/**
 * Computes 2 distinct teams (no character is reused).
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<string>} team1BuffExpressions - Buffs for slot 1.
 * @param {Array<string>} team2BuffExpressions - Buffs for slot 2.
 * @param {Array<object>} [buffOptions] - Selectable buff options.
 * @param {number} [maxOptions=50] - Max options to return.
 * @returns {Array<object>} A list of optimized teamPair objects.
 */
function computeBestDistinctTeamPairs(teams, team1BuffExpressions, team2BuffExpressions, buffOptions, maxOptions = 50) {
  return computeBestDistinctTeamGroups(teams, [team1BuffExpressions, team2BuffExpressions], buffOptions, maxOptions);
}

/**
 * Computes 3 distinct teams.
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<string>} team1BuffExpressions - Buffs for slot 1.
 * @param {Array<string>} team2BuffExpressions - Buffs for slot 2.
 * @param {Array<string>} team3BuffExpressions - Buffs for slot 3.
 * @param {Array<object>} [buffOptions] - Selectable buff options.
 * @param {number} [maxOptions=50] - Max options to return.
 * @returns {Array<object>} A list of optimized teamTriple objects.
 */
function computeBestDistinctTeamTriples(teams, team1BuffExpressions, team2BuffExpressions, team3BuffExpressions, buffOptions, maxOptions = 50) {
  return computeBestDistinctTeamGroups(
    teams,
    [team1BuffExpressions, team2BuffExpressions, team3BuffExpressions],
    buffOptions,
    maxOptions
  );
}

/**
 * Compute 4 distinct teams.
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<Array<string>>} buffExpressionsList - An array of 4 buff expression arrays.
 * @param {Array<object>} [buffOptions] - Selectable buff options.
 * @param {number} [maxOptions=50] - Max options to return.
 * @returns {Array<object>} A list of optimized 4-team objects.
 */
function computeBestDistinctTeamQuads(teams, buffExpressionsList, buffOptions, maxOptions = 50) {
  if (!buffExpressionsList || buffExpressionsList.length !== 4) {
    console.error("buffExpressionsList must be an array of 4 buff lists.");
    return [];
  }
  return computeBestDistinctTeamGroups(teams, buffExpressionsList, buffOptions, maxOptions);
}

/**
 * Computes 5 distinct teams.
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<Array<string>>} buffExpressionsList - An array of 5 buff expression arrays.
 * @param {Array<object>} [buffOptions] - Selectable buff options.
 * @param {number} [maxOptions=50] - Max options to return.
 * @returns {Array<object>} A list of optimized 5-team objects.
 */
function computeBestDistinctTeamQuints(teams, buffExpressionsList, buffOptions, maxOptions = 50) {
  if (!buffExpressionsList || buffExpressionsList.length !== 5) {
    console.error("buffExpressionsList must be an array of 5 buff lists.");
    return [];
  }
  return computeBestDistinctTeamGroups(teams, buffExpressionsList, buffOptions, maxOptions);
}

function hasUniqueCharacters(team1Chars, team2Chars) {
  return !team2Chars.includes(team1Chars[0]) && !team2Chars.includes(team1Chars[1]) && !team2Chars.includes(team1Chars[2]);
}

function all3HaveUniqueCharacters(team1Chars, team2Chars, team3Chars) {
  return hasUniqueCharacters(team1Chars, team2Chars) && hasUniqueCharacters(team1Chars, team3Chars) && hasUniqueCharacters(team2Chars, team3Chars);
}

/**
 * A generic helper to check uniqueness for any number of teams.
 * @param {Array<object>} teamList - An array of team objects.
 * @returns {boolean} True if all teams have unique characters from each other.
 */
function allTeamsHaveUniqueCharacters(teamList) {
  for (let i = 0; i < teamList.length; i++) {
    for (let j = i + 1; j < teamList.length; j++) {
      if (!hasUniqueCharacters(teamList[i].characters, teamList[j].characters)) {
        return false;
      }
    }
  }
  return true;
}