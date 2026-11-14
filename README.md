# xls-row-consolidator

A small, dependency-free **VBA toolkit for grouping, merging, flagging and
splitting rows in spreadsheet workbooks (for example Excel)**. Every macro is driven by a small
configuration block at the top of the module, so the same code works on any
table — logistics, inventory, finance, lab data, survey results, you name it.

> **Unofficial.** Not affiliated with, endorsed by, or sponsored by Microsoft
> (Excel). "xls" here is used only as a generic reference to spreadsheet
> workbooks.

## What it does

| Macro | Purpose |
|-------|---------|
| `ConsolidateRows` | Group rows by one or two key columns and write one summarized row per group to a new sheet (summing chosen numeric columns). |
| `MergeRowsByKey` | In-place: collapse duplicate-key rows into the first occurrence, summing numeric columns. |
| `MergeAdjacentMatchingRows` | Collapse runs of consecutive rows that share a key, summing numeric columns. |
| `FlagMatchedRows` | Highlight rows whose key also appears in a reference sheet. |
| `SplitTotalByRatio` | Spread a merged total across rows proportionally to a base column, reconciling rounding. |
| `AllocateIntoGroups` | (Advanced) partition a total into groups that respect optional average bounds; falls back to ratio split when bounds are off or the block is large. |
| `NormalizeText` / `SimplifyDelimitedField` / `IsChinese` / `ClassifyByKeyword` / `BubbleSortDescending` | Helper utilities for text cleanup and classification. |

All column letters, key columns, summed columns, sheet names, the highlight
color and the grouping bounds are **constants** — change them once and the
macros adapt to your data.

## Quick start

1. Open your workbook in Excel or a compatible spreadsheet app.
2. Press `Alt + F11` to open the VBA editor, then `File ▸ Import File…` and
   select `src/RowRollup.bas`.
3. Edit the `CONFIG` block at the top of the module to match your sheet:
   - `KEY_COL1` / `KEY_COL2` — the grouping column(s)
   - `SUM_COLS` — the numeric columns to aggregate
   - `RESULT_SHEET` / `REF_SHEET` — output and reference sheet names
4. Close the editor, press `Alt + F8`, and run the macro you need.

A ready-to-use example lives in `examples/sample-data.xlsx` (fictitious
`Group` / `Item` / `Pieces` / `Quantity` / `Weight` data) — open it and run
`ConsolidateRows` with the default config to see the result immediately.

See [`docs/Configuration.md`](docs/Configuration.md) for the full list of
settings and copy-paste recipes.

## Repository layout

```
xls-row-consolidator/
├── src/RowRollup.bas        # all macros + helpers (the only code file)
├── docs/Configuration.md    # how to point the macros at your data
├── examples/sample-data.xlsx# neutral fictitious sample workbook
├── LICENSE                  # MIT
├── PUSH_GUIDE.md            # how this repo was published to GitHub
├── .gitignore
└── .gitattributes
```

## Legal

This repository is a **neutral data-processing utility**. It contains no
business logic, no specific field vocabulary, no proprietary data and no
screenshots. The code is an independent, generic reimplementation intended to
be useful for ordinary spreadsheet work. The MIT license covers only the code
as published here.

Users are responsible for ensuring that any data they process and any use they
make of this tool comply with applicable laws and the terms of their own
workplace.

## License

Released under the [MIT License](LICENSE). Copyright (c) 2026 ivy92888-glitch.
