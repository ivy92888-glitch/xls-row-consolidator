# Configuration

All behavior is controlled by `Const` declarations at the top of
`src/RowRollup.bas`, inside the `CONFIG` block. You never need to touch the
logic — just edit these values.

| Constant | Default | Meaning |
|----------|---------|---------|
| `SRC_SHEET` | `""` | Sheet the data macros read. `""` = active sheet; otherwise the sheet name. |
| `KEY_COL1` | `"A"` | First grouping column (letter). |
| `KEY_COL2` | `"B"` | Second grouping column. Set to `""` for single-key grouping. |
| `SUM_COLS` | `"E,F,G"` | Comma-separated numeric columns to aggregate (sum). |
| `RESULT_SHEET` | `"Consolidated"` | Output sheet created by `ConsolidateRows`. |
| `REF_SHEET` | `"Reference"` | Sheet holding the reference key list for `FlagMatchedRows`. |
| `REF_KEY_COL1` | `"A"` | First key column of the reference sheet. |
| `REF_KEY_COL2` | `"B"` | Second key column of the reference sheet (`""` = single key). |
| `FLAG_COL` | `"B"` | Column highlighted when a key is matched. |
| `HILITE_COLOR` | `65535` | Fill color for highlights (yellow). |
| `TOTAL_COL` | `"E"` | Column holding the total to split (`SplitTotalByRatio` / `AllocateIntoGroups`). |
| `UNIT_COL` | `"F"` | Column used as the split-ratio base. |
| `WEIGHT_COL` | `"G"` | Column used by group allocation. |
| `GRP_MIN_AVG` | `0` | Min allowed average per group in `AllocateIntoGroups` (`0` = off). |
| `GRP_MAX_AVG` | `0` | Max allowed average per group (`0` = off). |
| `GRP_MAX_ITEMS` | `12` | Above this many rows in a block, allocation falls back to ratio split. |

## Recipes

### Summarize a sales export by region + product
```vb
KEY_COL1 = "C"   ' Region
KEY_COL2 = "D"   ' Product
SUM_COLS = "H,I" ' Units, Revenue
```

### Flatten duplicate ID rows, keeping the first
```vb
KEY_COL1 = "A"   ' ID
KEY_COL2 = ""    ' single-key mode
SUM_COLS = "E,F"
```
Run `MergeRowsByKey` on the active sheet.

### Flag rows already present in a master list
Put the master keys in a sheet named `Reference` (columns A, B), then run
`FlagMatchedRows` on your working sheet.

## Helper functions

These are plain functions you can call from your own code:

- `NormalizeText(text)` — trim, collapse spaces, strip bracketed content.
- `SimplifyDelimitedField(text, delim)` — keep one side of a `A/B` field.
- `IsChinese(text)` — heuristic CJK detection (useful for bilingual cleanup).
- `ClassifyByKeyword(map, text)` — classify `text` using a `Dictionary` you
  supply, e.g. `map("Plastic") = "pet,pp,abs"` maps any of those keywords to
  the class `"Plastic"`. No vocabulary is built in.
