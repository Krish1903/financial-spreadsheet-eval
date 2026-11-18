Audit the financial model below for three types of bugs:

- hardcode: a typed-in number where a formula should be
- off_by_one: a SUM or AVERAGE range that is one row or column too short or too long
- unit_mismatch: a value in the wrong unit (for example dollars vs thousands, or 25 vs 0.25)

The workbook is given as one line per non-empty cell: the cell address, then the formula and its computed value (formula -> value), or just the value if the cell has no formula.

Return only a JSON list of findings, with no other text:

[{"sheet": "...", "cell": "...", "bug_type": "hardcode | off_by_one | unit_mismatch", "explanation": "..."}]

Only report real bugs. If you find none, return [].

<workbook>
{WORKBOOK}
</workbook>
