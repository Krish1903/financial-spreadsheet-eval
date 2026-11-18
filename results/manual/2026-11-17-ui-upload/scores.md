# Manual run 0: xlsx uploaded through the chat UI (2026-11-17)

Setup: v0 base model (G&A 10%, growth 8%), 6-bug variant, audit_v1 prompt,
fresh incognito chat per model, file renamed to a neutral name.

All three models said the upload showed values only, no formulas and no cell addresses.
They guessed addresses by counting rows, so most cell references are off by 1 to 3 rows.
Scored by whether the right bug was found, not exact cell.

| Bug | Sonnet 5.5 | Opus 5.5 | Haiku 5.5 |
|---|---|---|---|
| 1. IS!E7 hardcode | miss | miss | miss |
| 2. CF!D7 hardcode | miss | miss | miss |
| 3. IS!F14 off_by_one | found, wrong type (hardcode) | found, wrong type (hardcode) | found |
| 4. Summary!B12 off_by_one | found | found | found |
| 5. Inputs!B12 unit_mismatch | found | found | found |
| 6. Summary!B5 unit_mismatch | found | found | found |
| Strict (right bug and type) | 3/6 | 3/6 | 4/6 |
| Location (right bug, any type) | 4/6 | 4/6 | 4/6 |
| False positives | 0 (IS!F10 flag is downstream, ignored) | 0 | 0 |
| Exact cell match | 2/6 | 0/6 | 0/6 |

## Takeaways
- Hardcodes were impossible to catch: with values only, a pasted number looks the same as a formula result.
- In v0, 2029 S&M + G&A (25% of revenue) exactly equaled 2028 EBITDA (27% = 25% x 1.08).
  Sonnet and Opus used that coincidence to wrongly call bug 3 a hardcode. Fixed in v1 inputs,
  and inject_manual.py now warns on any value collision.
- Sonnet's tax explanation says 100x; it is 25x. Explanations are not scored.
- Conclusion: UI upload is not a valid harness. Use serialize.py so the model sees addresses and formulas.
