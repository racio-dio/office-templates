# office-templates · declarative Excel template generator

[English](README.md) | [中文](README.zh-CN.md)

[![CI](https://github.com/racio-dio/office-templates/actions/workflows/ci.yml/badge.svg)](https://github.com/racio-dio/office-templates/actions/workflows/ci.yml)

Generate Excel workbooks that calculate as soon as you fill them in — formulas, dropdowns, conditional formatting and charts included. Pure Python + openpyxl, no Excel installation required.

Two ways to use it:

- **Ship the built-ins** — 5 ready-made office templates, parameterised with your company name, staff list and rates
- **Describe your own** — write a YAML file, get a workbook. No Python required

## Why

- **No code for new templates.** A spreadsheet is a schema: columns, types, formulas. Say it in YAML, get the `.xlsx`.
- **Templates you can diff.** YAML is reviewable and version-controllable; a binary `.xlsx` is not.
- **Parameterised output.** Generate the same layout for 40 people or 3, with your company name already in the title.

## Quick start

```bash
pip install -r requirements.txt
python build.py
```

Five `.xlsx` files land in `dist/`. Yellow cells are for you to fill in; everything else is formulas.

Install it as a command to use it anywhere:

```bash
pip install -e .
office-templates --list
office-templates -o ./my-templates
```

The CLI speaks English and Chinese: `--lang en` / `--lang zh`, or set `OFFICE_TEMPLATES_LANG`.

## CLI

```
python build.py [options]      # or: office-templates [options]
```

| Option | What it does |
|---|---|
| `-o, --out-dir <dir>` | output directory (default: `dist/`) |
| `-t, --only <...>` | build only these built-ins, by number or key |
| `--spec <file...>` | build custom templates from YAML/JSON definitions |
| `-l, --list` | list available templates and exit |
| `-q, --quiet` | only report errors |
| `-V, --version` | print version |
| `-c, --company <name>` etc. | template parameters, see below |
| `--lang zh\|en` | interface language |
| `-h, --help` | full help with examples |

### Template parameters

| Option | Effect |
|---|---|
| `-c, --company` | company name, prefixed to every title |
| `--staff` | staff names, comma separated |
| `--staff-file` | file of names, one per line (`#` = comment) |
| `--departments` | departments, comma separated |
| `--year` / `--month` | year / month used by the sheets |
| `--social-rate` | social insurance rate, e.g. `0.105` |
| `--fund-rate` | housing fund rate, e.g. `0.07` |
| `--start-date` | project start date `YYYY-MM-DD` (Gantt) |
| `--config` | JSON file with any of the above (CLI flags win) |

```bash
python build.py -c Acme --staff Alice,Bob,Carol --departments Sales,Engineering
python build.py --staff-file names.txt --year 2026 --month 11
python build.py --config examples/params.json
```

Row counts follow your staff list — 40 names produce 40 rows, no manual inserting.

## Built-in templates

| # | Template | Key | What it does |
|---|---|---|---|
| 1 | Attendance sheet | `kaoqin` | mark √ / late / leave / absent per day, attendance rate computed, weekends red |
| 2 | Payroll + payslips | `gongzi` | social insurance and housing fund deducted by rate, progressive income tax, printable payslips |
| 3 | Project Gantt | `gantt` | set start and duration, bars draw themselves, today marker, overdue flag |
| 4 | Personal budget | `jizhang` | category dropdowns, monthly totals, savings rate, spending pie chart |
| 5 | Inventory | `kucun` | log stock in/out, running balance, low-stock rows turn red |

> **Heads up:** these five ship with **Chinese** labels and follow Chinese payroll/tax conventions (progressive income tax, social insurance rates). The engine, CLI and YAML schema underneath are language-neutral — see Roadmap.

![Templates](screenshots/01_考勤.png)

## Define your own template in YAML

```bash
python build.py --spec examples/specs/expense-report.yaml
python build.py --spec my.yaml -c Acme -o ./out
```

```yaml
filename: expense-report.xlsx
title: Expense report
subtitle: Fill the yellow cells; anything over 1000 is flagged red

sheets:
  - name: Items
    header_row: 3
    rows: 30
    freeze: A4
    totals: true
    columns:
      - {label: Date, width: 12, type: date, format: yyyy-mm-dd, fill: FFF9E5}
      - {label: Category, width: 12, choices: [Travel, Meals, Lodging, Supplies, Other]}
      - {label: Amount, width: 12, type: number, format: "#,##0.00", total: true, sample: [553, 680]}
      - {label: Note, width: 24}
    conditionals:
      - range: "A{r1}:D{last}"
        when: "$C{r1}>1000"
        font_color: C00000
        bold: true

  - name: By category
    header_row: 3
    rows: 5
    totals: true
    columns:
      - {label: Category, width: 14, sample: [Travel, Meals, Lodging, Supplies, Other]}
      - {label: Amount, width: 14, format: "#,##0.00", total: true,
         formula: "=SUMIFS(Items!$C:$C,Items!$B:$B,A{r})"}
      - {label: Share, width: 10, format: "0.0%", formula: "=IFERROR(B{r}/$B${total},0)"}
```

### Fields

| Level | Field | Meaning |
|---|---|---|
| root | `filename` | output name, must end with `.xlsx` |
| | `title` / `subtitle` | heading and subheading (overridable per sheet) |
| | `sheets` | list of sheets, at least one |
| sheet | `name` / `columns` | sheet name and its columns (required) |
| | `header_row` / `rows` | header row (default 3), data rows (default 30) |
| | `freeze` | freeze panes, e.g. `A4` |
| | `totals` | add a totals row, summing columns marked `total: true` |
| | `conditionals` | conditional formatting rules |
| column | `label` | column heading (required) |
| | `type` | `text` (default) / `number` / `date` |
| | `width` / `format` | column width, number format |
| | `formula` | formula with placeholders (used instead of `sample`) |
| | `choices` | dropdown options |
| | `sample` | example values, filled row by row |
| | `fill` | background colour (`FFF9E5` marks an input area) |
| | `total` | include in the totals row |
| conditional | `range` / `when` | target range and condition, placeholders allowed |
| | `font_color` / `fill` / `bold` | styling when it matches |

### Placeholders

Usable in `formula`, `range` and `when`:

| Placeholder | Resolves to |
|---|---|
| `{r}` | current data row |
| `{r1}` | first data row |
| `{last}` | last data row |
| `{total}` | totals row (same as `{last}` when `totals` is off) |

Bad definitions fail loudly with a specific message and exit code 2.

## Project structure

```
office-templates/
├── build.py                    # CLI entry point
├── office_templates/
│   ├── options.py              # parameters (company, staff, rates…)
│   ├── spec.py                 # YAML/JSON definition → dataclasses
│   ├── render.py               # definition → Workbook
│   ├── i18n.py                 # zh/en interface strings
│   ├── styles.py               # colours, title/header/body helpers
│   ├── registry.py             # built-in registry, selection, saving
│   ├── cli.py                  # argparse CLI
│   └── templates/              # built-ins, one module each
├── examples/
│   ├── params.json             # parameter file example
│   ├── staff.txt               # name list example
│   └── specs/                  # YAML template definitions
├── tests/                      # smoke + spec tests
└── .github/workflows/ci.yml    # CI on Python 3.9 / 3.11 / 3.13
```

Use it as a library:

```python
from office_templates import Options, build_templates, resolve

opts = Options(company="Acme", staff=("Alice", "Bob"), month=11)
build_templates(resolve(["kaoqin", "gongzi"]), "./out", opts)
```

```python
from office_templates import load_spec, render_spec

wb = render_spec(load_spec("my.yaml"), Options(company="Acme"))
wb.save("out/my-template.xlsx")
```

## Roadmap

- [x] CLI, parameters, YAML-defined templates, CI
- [ ] Locale packs for the built-in templates (`zh-CN`, `en-US`): fonts, date formats, labels
- [ ] English-first built-ins — expense report, timesheet, inventory — with Western conventions

## Development

```bash
python -m unittest discover -s tests -v
```

Tests cover: every built-in generates and reopens cleanly, parameter validation, parameters actually landing in cells, and YAML definitions rendering correctly. CI runs them on Python 3.9 / 3.11 / 3.13 and uploads the generated templates as build artifacts.

## Notes

- Python 3.9+. Runtime deps: `openpyxl` (writing xlsx) and `PyYAML` (YAML definitions).
- The built-in payroll sheet uses a simplified monthly income tax (5000 threshold) — fine for a template, not for actual payroll.

## License

MIT
