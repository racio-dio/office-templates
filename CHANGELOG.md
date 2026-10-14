# Changelog

## 0.5.1 (2026-10-15)

Fixes
- A typo in a placeholder (`{row}` instead of `{r}`) used to crash with a KeyError traceback; specs are now checked when loaded and the error names the placeholder.
- Duplicate sheet names were silently renamed (`Items` became `Items1`), which broke formulas pointing at them. They are now an error, compared case-insensitively like Excel does.
- Sheet names over 31 characters produced files Excel offers to "repair". Rejected now, along with `\ / ? * [ ] :`.
- `header_row` 1 or 2 could overwrite the title or subtitle. Now validated.
- Dropdown choices containing a comma, or adding up to more than 255 characters, produced broken validation lists. Rejected with a clear message.
- `filename` could contain a path (`../x.xlsx`) and write outside the output folder. Must be a plain file name now.
- Numbers given as strings in YAML/JSON (`rows: "10"`, `"month": "11"`) crashed; they are converted, and non-numbers give a readable error.
- Default year, month and Gantt start date were hard-coded to October 2026. They now default to today.
- Render errors in the CLI exit with code 2 instead of a traceback.

Changes
- English is the default interface language when the environment doesn't say otherwise.
- Specs written in English use Calibri and label the totals row `Total`; Chinese specs keep 微软雅黑 and `合计`. Override the label with `total_label`.
- Every generated sheet is set to print one page wide (landscape when it has more than 7 columns).
- The English example spec uses US formats and sample data.
- README rewritten with a quick start, a screenshot and a short demo.
- New tests for all of the above (`tests/test_validation.py`).
