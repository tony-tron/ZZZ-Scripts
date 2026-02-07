/** @OnlyCurrentDoc */

function getAllTeams(minTeamStrength) {
  var data = getSortedTeamsSheet().getDataRange().getValues();
  var teams = [];
  for (var r = 1; r < data.length; r++) {
    var team = {
      characters : [data[r][0], data[r][1], data[r][2]],
      strength : data[r][3],
    }
    if (team.strength < minTeamStrength) continue;
    addBuffParamsToTeam(team);
    teams.push(team);
  }
  return teams;
}

/**
 * Computes 1 distinct team, for completeness (all teams are distinct so why would you want this?).
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<string>} team1BuffExpressions - Buffs for slot 1.
 * @returns {Array<object>} A list of optimized 1-team objects.
 */
function computeBestDistinctTeamSingles(teams, team1BuffExpressions, buffOptions) {
  const teamSingles = [];
  const buffList = [team1BuffExpressions];

  for (let i = 0; i < teams.length; i++) {
    const team = teams[i];
    
    // k=1, no uniqueness check needed
    const teamSlots = [{ team }];
    const result = optimizeTeamAssignments(teamSlots, buffList);

    const teamSingle = {
      team1: result.optimizedSlots[0].team,
      team1Bonus: result.bonuses[0],
      team1ChosenBonusName: "",
      team1ChosenBonus: 0,
      totalStrength: function() { return Math.round((this.team1.strength + this.team1Bonus + this.team1ChosenBonus) * 1000) / 1000 },
      minStrength: function() { return this.totalStrength() }, // Only one team
    };

    // Apply "chosen" buff
    var buffOption = chooseBestBuffForTeam(teamSingle.team1, buffOptions);
    if (buffOption != undefined) {
      teamSingle.team1ChosenBonusName = buffOption.name;
      teamSingle.team1ChosenBonus = buffOption.bonus;
    }
    
    teamSingles.push(teamSingle);
  }
  return teamSingles;
}

/**
 * Computes 2 distinct teams (no character is reused).
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<string>} team1BuffExpressions - Buffs for slot 1.
 * @param {Array<string>} team2BuffExpressions - Buffs for slot 2.
 * @returns {Array<object>} A list of optimized teamPair objects.
 */
function computeBestDistinctTeamPairs(teams, team1BuffExpressions, team2BuffExpressions, buffOptions) {
  const teamPairs = [];
  const buffList = [team1BuffExpressions, team2BuffExpressions]; // Fixed buff slots

  for (let i = 0; i < teams.length; i++) {
    for (let j = i + 1; j < teams.length; j++) {
      const team1 = teams[i];
      const team2 = teams[j];

      if (hasUniqueCharacters(team1.characters, team2.characters)) {
        // We just pass the team object. The optimizer will use team.team
        const teamSlots = [{ team: team1, originalIndex: i }, { team: team2, originalIndex: j }];
        
        // Run the generic optimization
        const result = optimizeTeamAssignments(teamSlots, buffList);

        // Create the teamPair object from the *optimized* result
        const teamPair = {
          team1: result.optimizedSlots[0].team, // Team now in slot 1
          team2: result.optimizedSlots[1].team, // Team now in slot 2
          team1Bonus: result.bonuses[0],
          team2Bonus: result.bonuses[1],
          totalStrength: function() { return Math.round((this.team1.strength + this.team2.strength + this.team1Bonus + this.team2Bonus) * 1000) / 1000 },
          minStrength: function() { return Math.round(Math.min(this.team1.strength + this.team1Bonus, this.team2.strength + this.team2Bonus) * 1000) / 1000 },
        };

        // Apply "chosen" buffs
        [teamPair.team1, teamPair.team2].forEach((team, idx) => {
          var buffOption = chooseBestBuffForTeam(team, buffOptions);
          if (buffOption != undefined) {
            teamPair[`team${idx+1}ChosenBonusName`] = buffOption.name;
            teamPair[`team${idx+1}ChosenBonus`] = buffOption.bonus;
          }
        });

        teamPairs.push(teamPair);
      }
    }
  }
  return teamPairs;
}

/**
 * Computes 3 distinct teams.
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<string>} team1BuffExpressions - Buffs for slot 1.
 * @param {Array<string>} team2BuffExpressions - Buffs for slot 2.
 * @param {Array<string>} team3BuffExpressions - Buffs for slot 3.
 * @returns {Array<object>} A list of optimized teamTriple objects.
 */
function computeBestDistinctTeamTriples(teams, team1BuffExpressions, team2BuffExpressions, team3BuffExpressions, buffOptions) {
  const teamTriples = [];
  const buffList = [team1BuffExpressions, team2BuffExpressions, team3BuffExpressions]; // Fixed buff slots

  for (let i = 0; i < teams.length; i++) {
    for (let j = i + 1; j < teams.length; j++) {
      for (let k = j + 1; k < teams.length; k++) {
        const team1 = teams[i];
        const team2 = teams[j];
        const team3 = teams[k];

        if (all3HaveUniqueCharacters(team1.characters, team2.characters, team3.characters)) {
          
          const teamSlots = [{ team: team1 }, { team: team2 }, { team: team3 }];
          
          // Run optimization for the 3 slots
          const result = optimizeTeamAssignments(teamSlots, buffList);

          // Create the triple with the *optimized* team order
          const teamTriple = {
            team1: result.optimizedSlots[0].team, // Team in slot 1
            team2: result.optimizedSlots[1].team, // Team in slot 2
            team3: result.optimizedSlots[2].team, // Team in slot 3
            team1Bonus: result.bonuses[0],
            team2Bonus: result.bonuses[1],
            team3Bonus: result.bonuses[2],
            team1ChosenBonusName: "",
            team2ChosenBonusName: "",
            team3ChosenBonusName: "",
            team1ChosenBonus: 0,
            team2ChosenBonus: 0,
            team3ChosenBonus: 0,
            totalStrength: function() { return Math.round((this.team1.strength + this.team2.strength + this.team3.strength + this.team1Bonus + this.team2Bonus + this.team3Bonus + this.team1ChosenBonus + this.team2ChosenBonus + this.team3ChosenBonus) * 1000) / 1000 },
            minStrength: function() { return Math.round(Math.min(this.team1.strength + this.team1Bonus + this.team1ChosenBonus, this.team2.strength + this.team2Bonus + this.team2ChosenBonus, this.team3.strength + this.team3Bonus + this.team3ChosenBonus) * 1000) / 1000 },
          };

          // Apply the separate "chosen" buffs (from your original logic)
          // This is applied to the team *after* it has been assigned to its optimal slot
          var buffOption = chooseBestBuffForTeam(teamTriple.team1, buffOptions);
          if (buffOption != undefined) {
            teamTriple.team1ChosenBonusName = buffOption.name;
            teamTriple.team1ChosenBonus = buffOption.bonus;
          }
          buffOption = chooseBestBuffForTeam(teamTriple.team2, buffOptions);
          if (buffOption != undefined) {
            teamTriple.team2ChosenBonusName = buffOption.name;
            teamTriple.team2ChosenBonus = buffOption.bonus;
          }
          buffOption = chooseBestBuffForTeam(teamTriple.team3, buffOptions);
          if (buffOption != undefined) {
            teamTriple.team3ChosenBonusName = buffOption.name;
            teamTriple.team3ChosenBonus = buffOption.bonus;
          }

          teamTriples.push(teamTriple);
        }
      }
    }
  }
  return teamTriples;
}

/**
 * Compute 4 distinct teams.
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<Array<string>>} buffExpressionsList - An array of 4 buff expression arrays.
 * @returns {Array<object>} A list of optimized 4-team objects.
 */
function computeBestDistinctTeamQuads(teams, buffExpressionsList, buffOptions) {
  if (!buffExpressionsList || buffExpressionsList.length !== 4) {
      console.error("buffExpressionsList must be an array of 4 buff lists.");
      return [];
  }

  const teamQuads = [];
  
  for (let i = 0; i < teams.length; i++) {
    for (let j = i + 1; j < teams.length; j++) {
      for (let k = j + 1; k < teams.length; k++) {
        for (let l = k + 1; l < teams.length; l++) {
          const teamList = [teams[i], teams[j], teams[k], teams[l]];

          if (allTeamsHaveUniqueCharacters(teamList)) {
            const teamSlots = teamList.map(team => ({ team }));
            const result = optimizeTeamAssignments(teamSlots, buffExpressionsList);

            const teamQuad = {
              team1: result.optimizedSlots[0].team,
              team2: result.optimizedSlots[1].team,
              team3: result.optimizedSlots[2].team,
              team4: result.optimizedSlots[3].team,
              team1Bonus: result.bonuses[0],
              team2Bonus: result.bonuses[1],
              team3Bonus: result.bonuses[2],
              team4Bonus: result.bonuses[3],
              team1ChosenBonusName: "",
              team2ChosenBonusName: "",
              team3ChosenBonusName: "",
              team4ChosenBonusName: "",
              team1ChosenBonus: 0,
              team2ChosenBonus: 0,
              team3ChosenBonus: 0,
              team4ChosenBonus: 0,
            };

            teamQuad.totalStrength = function() {
              let base = this.team1.strength + this.team2.strength + this.team3.strength + this.team4.strength;
              let slotBonus = this.team1Bonus + this.team2Bonus + this.team3Bonus + this.team4Bonus;
              let chosenBonus = this.team1ChosenBonus + this.team2ChosenBonus + this.team3ChosenBonus + this.team4ChosenBonus;
              return Math.round((base + slotBonus + chosenBonus) * 1000) / 1000;
            };
            teamQuad.minStrength = function() {
              return Math.round(Math.min(
                this.team1.strength + this.team1Bonus + this.team1ChosenBonus,
                this.team2.strength + this.team2Bonus + this.team2ChosenBonus,
                this.team3.strength + this.team3Bonus + this.team3ChosenBonus,
                this.team4.strength + this.team4Bonus + this.team4ChosenBonus
              ) * 1000) / 1000;
            };

            // Apply "chosen" buffs
            [teamQuad.team1, teamQuad.team2, teamQuad.team3, teamQuad.team4].forEach((team, idx) => {
              var buffOption = chooseBestBuffForTeam(team, buffOptions);
              if (buffOption != undefined) {
                teamQuad[`team${idx+1}ChosenBonusName`] = buffOption.name;
                teamQuad[`team${idx+1}ChosenBonus`] = buffOption.bonus;
              }
            });
            
            teamQuads.push(teamQuad);
          }
        }
      }
    }
  }
  return teamQuads;
}

/**
 * Computes 5 distinct teams.
 *
 * @param {Array<object>} teams - The list of all available teams.
 * @param {Array<Array<string>>} buffExpressionsList - An array of 5 buff expression arrays.
 * @returns {Array<object>} A list of optimized 5-team objects.
 */
function computeBestDistinctTeamQuints(teams, buffExpressionsList, buffOptions) {
  if (!buffExpressionsList || buffExpressionsList.length !== 5) {
      console.error("buffExpressionsList must be an array of 5 buff lists.");
      return [];
  }

  const teamQuints = [];
  const TARGET_DEPTH = 5;

  function findTeamCombinations(startIndex, currentTeams, usedCharacters) {
    // Base case: we have 5 teams
    if (currentTeams.length === TARGET_DEPTH) {

      const teamSlots = currentTeams.map(team => ({ team }));

      // Run optimization for the 5 slots
      const result = optimizeTeamAssignments(teamSlots, buffExpressionsList);

      // Create the quintuple with the *optimized* team order
      const teamQuint = {
        team1: result.optimizedSlots[0].team,
        team2: result.optimizedSlots[1].team,
        team3: result.optimizedSlots[2].team,
        team4: result.optimizedSlots[3].team,
        team5: result.optimizedSlots[4].team,
        team1Bonus: result.bonuses[0],
        team2Bonus: result.bonuses[1],
        team3Bonus: result.bonuses[2],
        team4Bonus: result.bonuses[3],
        team5Bonus: result.bonuses[4],
        team1ChosenBonusName: "",
        team2ChosenBonusName: "",
        team3ChosenBonusName: "",
        team4ChosenBonusName: "",
        team5ChosenBonusName: "",
        team1ChosenBonus: 0,
        team2ChosenBonus: 0,
        team3ChosenBonus: 0,
        team4ChosenBonus: 0,
        team5ChosenBonus: 0,
      };

      // Add totalStrength and minStrength functions
      teamQuint.totalStrength = function() {
        let base = this.team1.strength + this.team2.strength + this.team3.strength + this.team4.strength + this.team5.strength;
        let bonus = this.team1Bonus + this.team2Bonus + this.team3Bonus + this.team4Bonus + this.team5Bonus;
        let chosenBonus = this.team1ChosenBonus + this.team2ChosenBonus + this.team3ChosenBonus + this.team4ChosenBonus + this.team5ChosenBonus;
        return Math.round((base + bonus + chosenBonus) * 1000) / 1000;
      };

      teamQuint.minStrength = function() {
        return Math.round(Math.min(
          this.team1.strength + this.team1Bonus + this.team1ChosenBonus,
          this.team2.strength + this.team2Bonus + this.team2ChosenBonus,
          this.team3.strength + this.team3Bonus + this.team3ChosenBonus,
          this.team4.strength + this.team4Bonus + this.team4ChosenBonus,
          this.team5.strength + this.team5Bonus + this.team5ChosenBonus
        ) * 1000) / 1000;
      };

      // Apply "chosen" buffs
      [teamQuint.team1, teamQuint.team2, teamQuint.team3, teamQuint.team4, teamQuint.team5].forEach((team, idx) => {
        var buffOption = chooseBestBuffForTeam(team, buffOptions);
        if (buffOption != undefined) {
          teamQuint[`team${idx+1}ChosenBonusName`] = buffOption.name;
          teamQuint[`team${idx+1}ChosenBonus`] = buffOption.bonus;
        }
      });

      teamQuints.push(teamQuint);
      return;
    }

    // Recursive step with pruning
    for (let i = startIndex; i < teams.length; i++) {
      const team = teams[i];
      const chars = team.characters;

      // Check for character overlap (Pruning)
      let overlap = false;
      for (let c = 0; c < chars.length; c++) {
        if (usedCharacters.has(chars[c])) {
          overlap = true;
          break;
        }
      }
      if (overlap) continue;

      // Add characters to set
      for (let c = 0; c < chars.length; c++) {
        usedCharacters.add(chars[c]);
      }

      // Recurse
      findTeamCombinations(i + 1, currentTeams.concat([team]), usedCharacters);

      // Backtrack: remove characters from set
      for (let c = 0; c < chars.length; c++) {
        usedCharacters.delete(chars[c]);
      }
    }
  }

  // Start recursion
  findTeamCombinations(0, [], new Set());

  return teamQuints;
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