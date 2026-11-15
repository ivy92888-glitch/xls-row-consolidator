# xls-row-consolidator / 表格行处理工具箱

> 零依赖的 VBA 工具箱：对电子表格里的「行」做分组、合并、标记、拆分——按关键字归组求和、合并重复行、标记命中、按比例拆分合计、按约束分配。同一个配置块驱动，套到任何表都能用。
> A dependency-free VBA toolkit for grouping, merging, flagging, and splitting rows in spreadsheet workbooks. One configuration block drives every macro.

> **非官方。** 与 Microsoft（Excel）无隶属、赞助或背书关系；"xls" 仅作为电子表格工作簿的通用指代。

---

## 痛点：它解决什么问题（中文）

Excel 里大量日常工作，本质都是对「行」的搬运与重组，而手工做这些事又慢又容易错：

- **按关键字归组求和**：把几十上百行按「组别 / 品名」合并成一行汇总（如按客户汇总金额）；
- **合并重复行**：同一关键信息出现多次，要把它们并到第一条、并把数值列加总；
- **标记命中行**：找出「在某个参考表里也出现过的行」并高亮（如对账差异、黑名单筛查）；
- **按比例拆分合计**：把一个合并后的总数，按某基准列（如重量、金额）分摊回各行，还要处理四舍五入造成的尾差；
- **按约束分配**：把一个总量拆进若干组、尽量满足「每组均值在某个区间」。

**手工做法的崩溃点**：筛选、排序、SUMIF、复制粘贴、手填拆分比例……步骤多、易漏、改一次需求就要重来一遍。更麻烦的是这些操作往往散落在不同宏或人工步骤里，没有统一入口。

**这个工具箱的价值**：把上面这些高频「行处理」动作，封装成一组**配置驱动**的宏——列号、关键列、求和列、表名、高亮色、分组区间全都是模块顶部的常量，改一次就能套到你的表。非程序员也能直接用，不用重写逻辑。

## What problem it solves (English)

Most spreadsheet busywork is really "row gymnastics" — group-and-sum, merge duplicates, flag rows that appear in a reference list, split a total back across rows by ratio, or allocate a total under bounds. Doing these by hand is slow and error-prone. This toolkit packages them as a set of configuration-driven macros; change a few constants at the top and they adapt to any table.

---

## 快速运行（中文）

**环境要求**

- Excel 或任何兼容的电子表格软件（支持 VBA）；
- 一张待处理的数据表。

**三步跑起来**

1. 用 Excel（或兼容软件）打开工作簿，按 `Alt + F11` 打开 VBA 编辑器；
2. `文件 ▸ 导入文件…`，选择 `src/RowRollup.bas`；
3. 改 `RowRollup.bas` 顶部的 `CONFIG` 配置块（关键列、求和列、输出/参考表名），关闭编辑器，按 `Alt + F8` 运行你需要的宏。

**想立刻看到效果**：打开 `examples/sample-data.xlsx`（虚构的 `Group / Item / Pieces / Quantity / Weight` 数据），用默认配置运行 `ConsolidateRows` 即可看到汇总结果。

**包含的宏**（详见文件内注释）

| 宏 | 作用 |
|----|------|
| `ConsolidateRows` | 按 1~2 个关键列归组，每组写一行汇总（求指定数值列之和）到新表 |
| `MergeRowsByKey` | 原地合并重复关键行，数值列求和 |
| `MergeAdjacentMatchingRows` | 合并连续相同关键字的行块，数值列求和 |
| `FlagMatchedRows` | 高亮「关键列也出现在参考表」的行 |
| `SplitTotalByRatio` | 按基准列把合并总数按比例摊回各行，并校正尾差 |
| `AllocateIntoGroups` | （进阶）把总量拆进各组并尽量满足均值区间；越界或大块时回退按比例拆分 |

## Quick start (English)

1. Open your workbook in Excel, press `Alt+F11`, `File ▸ Import File…` → `src/RowRollup.bas`.
2. Edit the `CONFIG` block at the top (key columns, summed columns, sheet names).
3. Press `Alt+F8` and run the macro you need. Try `examples/sample-data.xlsx` with `ConsolidateRows` first.

---

## 工作原理 / How it works

所有宏都由模块顶部的 `CONFIG` 常量驱动：列字母、关键列、求和列、表名、高亮色、分组边界都是常量，改一次即可适配你的数据。文本清理与分类还附带 `NormalizeText / SimplifyDelimitedField / IsChinese / ClassifyByKeyword / BubbleSortDescending` 等辅助函数。完整参数与复制即用示例见 [`docs/Configuration.md`](docs/Configuration.md)。

## 配置要点 / Configuration

列号、关键列、求和列、输出/参考表名、高亮色、分组边界均为常量。**参数名唯一真源是 `src/RowRollup.bas` 中的 `CONFIG` 块**，修改代码后请同步此处。

## 适用场景 / Use cases

物流、库存、财务、实验数据、调研结果——任何需要对「行」做归组、合并、标记、拆分、分配的表格工作。

## 法律声明 / Legal

本仓库是一个**中立的数据处理工具**，不含业务逻辑、特定字段词汇、专有数据或截图。代码为独立、通用的重写，旨在服务日常表格工作。MIT 许可证仅覆盖此处发布的代码。使用者需确保所处理的数据及用途符合所在地法律与所在单位政策。

## License

[MIT License](LICENSE) · Copyright (c) 2026 ivy92888-glitch.

---

<details><summary>维护说明（中英双语单文件，怎么避免漂移）</summary>

本文档是**单文件中英双语**：中文为权威详细描述，英文为简洁摘要（不强求逐句对等）。所有参数名（如 `KEY_COL1`、`SUM_COLS`）直接取自源码常量，唯一真源是 `src/RowRollup.bas` 中的 `CONFIG` 块。修改代码逻辑时，只要同步更新源码常量名与本文档「配置要点」中的同名常量即可，叙述性文字按需增删，不影响一致性。

</details>
