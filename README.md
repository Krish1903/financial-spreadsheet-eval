# financial-spreadsheet-eval

Measures how well LLM agents find bugs in financial models.

## How it works
1. Build a clean model (data/clean)
2. Inject known bugs and record them in an answer key (data/truth)
3. Serialize the bugged workbook to text (data/serialized) and give it to the agent
4. Score the agent's findings against the answer key

The agent only ever sees files in data/serialized. Never data/truth or data/clean.

## Bug types (v1)
- hardcode: a formula replaced by a typed number. Any constant outside the Inputs sheet is a bug.
- off_by_one: a SUM/AVERAGE range one row or column too short or too long
- unit_mismatch: values in the wrong unit (dollars vs thousands, 25 vs 0.25)

## Finding format
{"sheet": "IS", "cell": "E7", "bug_type": "hardcode", "explanation": "..."}

## Scoring
- Strict match: right cell and right bug type
- Location match: right cell only
- Flags on cells downstream of a bug are ignored
- Metrics: precision, recall (overall and per bug type), false positives on clean files, cost, latency

## Setup
Requires Python 3.12, uv, and LibreOffice (Mac: brew install --cask libreoffice).

    uv add openpyxl

## Commands
    uv run python -m src.auditor.generate        # build data/clean/fin-eval.xlsx
    uv run python -m src.auditor.inject_manual   # 6-bug variant + answer key
    uv run python -m src.auditor.serialize data/injected/fin-eval-bugs.xlsx data/serialized/fin-eval-bugs.txt

## Results so far
- results/manual/2026-11-17-ui-upload: chat UI upload, values only. Not a valid harness, see scores.md.
