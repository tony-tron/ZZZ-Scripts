1. **Update `02_Characters.js` to add new columns:**
   - Modify `initCharactersColumns` to add `sharpnessRegenBenefit` between `energyRegenBenefit` and `ultimateFocus`. This will require incrementing the column indices of all subsequent columns by 1.
   - Add `gashBuildup` at the very end of the `characterColumns` object with the appropriate column index.
2. **Update `09_BuffUtils.js` to include the new stats:**
   - In `initCharsToBuffParams`, add `sharpDamageBonusBenefit` calculated similarly to `lacerationDamageBonusBenefit` but with a 1.5x multiplier instead of 3x.
   - Add parsing for `sharpnessRegenBenefit` and `gashBuildup` from the `charactersData` using the new columns.
   - In `Team.prototype.initStats`, aggregate `SharpDamageBonusBenefit`, `SharpnessRegenBenefit`, and `GashBuildup` for the whole team.
3. **Verify syntax with `node -c`:**
   - Run `node -c 02_Characters.js` and `node -c 09_BuffUtils.js` to ensure there are no syntax errors introduced.
4. **Complete pre-commit steps to ensure proper testing, verification, review, and reflection are done.**
5. **Submit changes.**
   - Push code.
