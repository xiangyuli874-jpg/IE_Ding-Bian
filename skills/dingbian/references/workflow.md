# 完整定编流程

## 命令约定

在 `E:\AI\dingbian` 中运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage <阶段>
```

每次从命令输出读取新的绝对文件路径，将它作为下一阶段的 `<当前文件>`。不要继续使用原始文件。

每个阶段结束后运行：

```powershell
python "E:\AI\dingbian\skills\dingbian\scripts\inspect_workbook.py" "<当前文件>"
```

根据 `next_action` 继续：`resolve_foundation_data` 表示先补基础数据；`run_standard_format_and_decomposition` 表示进入标准台数、格式和分解；`run_decompose_extra_summary` 表示额外订单信息汇总缺少最新产能规划指标；`run_classify` 表示分解结果已具备但缺少最终分类；`review_unclassified_data` 表示分类已执行但仍有未分类行，需要向用户说明并询问是否继续补规则。

注意：不要凭文件名判断阶段完成。最新项目结果中可能出现文件名包含“分类结果”，但工作簿仍缺少“分类结果汇总表/未分类数据”的情况；必须以检查器输出为准。

## 1. 基础数据异常并行准备

1. 优先运行 `prepare-foundation-data`，一次性完成：
   - 删除 `订单数` 为空的行；
   - 删除 `物料编码` 为 `Z4U6010100` 的订单行；
   - 删除行后刷新 `系数` 和 `钣金型号` 的 VLOOKUP 当前行引用，避免公式仍查旧行号；
   - 摘取 `系数` 异常到 `系数补充`；
   - 摘取 `钣金型号` 异常到 `钣金型号补充`，异常包含 `#N/A`、`N/A`、空白和 `0`；
   - 摘取 `物料描述` 异常到 `物料描述补充`。

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage prepare-foundation-data
```

2. 若需要分阶段重试，也可以先运行 `cleanup-order-rows`，再按 `prepare-coefficients`、`prepare-sheet-metal`、`prepare-material-description` 连续执行。不要因为前一项仍有待补就停止后两项的摘取。
3. 运行 `inspect_workbook.py` 检查三项待补数量。
4. 若 `coefficient_pending`、`sheet_metal_pending`、`material_description_pending` 任一大于 0，暂停进入主表格式、标准台数和分解阶段，先按下面方式回填。

防错原则：所有基础数据补充都优先采用“补充表确认机制”。也就是外部查询表或 BOM 能匹配出的候选值，先写入对应的 `系数补充`、`钣金型号补充` 或 `物料描述补充`，并在关键字段前增加匹配说明/来源列供人工确认；确认后再通过对应 `apply-manual-*` 阶段写回主表。只有用户明确要求“直接回填主表”时，才使用直接 `fill-*` 阶段。若已经生成分解结果后又发现基础数据仍有待补，不要把该文件视为完成；补齐后从受影响的下游阶段继续重跑到 `classify`。

## 2. 系数回填

若缺失且用户提供了唯一系数查询表，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage fill-coefficients --coefficient-lookup "<系数查询表>"
```

若没有查询表，暂停并索取。查询后仍缺失，要求用户填写当前结果中的“系数补充”或“系数仍缺失”，保存并关闭 Excel。用户返回后运行 `apply-manual-coefficients`。

若用户确认“系数补充”中的剩余编码是无需处理/可删除的订单，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage cleanup-coefficient-supplement-rows
```

## 3. 钣金型号回填

若缺失且用户提供了唯一钣金型号查询表，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage fill-sheet-metal --sheet-metal-lookup "<钣金型号查询表>"
```

没有查询表时暂停索取。查询后仍缺失，要求用户填写“钣金型号补充”。用户返回后运行 `apply-manual-sheet-metal`。

若用户提供的是 BOM 表，且表内包含 `BOM号`、`物料描述`，需要按主表当前“钣金型号缺失/异常”的 `物料编码` 匹配 `BOM号`，再取同一 BOM 下 `物料描述` 含 `箱体组件` 的行回填钣金型号，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage suggest-sheet-metal-bom --sheet-metal-bom-lookup "<钣金型号BOM表>"
```

该阶段只把 BOM 匹配到的箱体组件候选写入“钣金型号补充”，不直接写主表。用户确认候选值正确、必要时手工修改“钣金型号补充”后，再运行 `apply-manual-sheet-metal` 写回主表。若用户明确要求跳过确认，也可以运行 `fill-sheet-metal-bom` 直接写回主表。

BOM 候选后仍缺失，再刷新“钣金型号补充”供用户人工补充。

注意：`箱体组件` BOM 表通常是同事按“钣金型号补充”里的缺失物料编码从系统导出的整机 BOM，里面可能同时包含滚筒和波轮。BOM 回填不能按线体一刀切，也不能覆盖主表已有正常钣金型号；只能回填当前仍为 `#N/A`、`N/A`、空白或 `0` 的行，并要防止明显滚筒/波轮组件错配。

若用户确认“物料描述补充”中的剩余编码都是未上单订单，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage cleanup-missing-material-description-rows
```

该阶段会删除主表中当前物料描述为空白或 `#N/A/N/A` 的订单行，并刷新系数和钣金型号公式行引用。

## 4. 物料描述回填

若缺失且用户提供了唯一物料描述查询表，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage fill-material-description --material-description-lookup "<物料描述查询表>"
```

没有查询表时暂停索取。查询后仍缺失，要求用户填写“物料描述补充”或“物料描述仍缺失”。用户返回后运行 `apply-manual-material-description`。

## 5. 主表格式、标准台数、完整排单分解与最终分类

基础数据三项全部补齐后，按以下固定顺序逐个运行；每个阶段必须使用上一阶段的新输出：

1. `prepare-standard-units`
2. `format-main-sheet`
3. `decompose-skd`
4. `decompose-rolling-remarks`
5. `decompose-t7p7t5p5-dreame`
6. `decompose-t9p9`
7. `decompose-t9p9-dryer`
8. `decompose-t10p10`
9. `decompose-c6-heat-pump-dryer`
10. `decompose-composite-penguin-c6`
11. `decompose-rolling-final`
12. `decompose-wave-basic`
13. `decompose-wave-final`
14. `decompose-extra-summary`
15. `classify`

不要把 `decompose-extra-summary` 提前到完整分解之前。
`classify` 是当前定编默认必跑的最终阶段。没有生成“分类结果汇总表”和“未分类数据”时，不要称为定编完成。若用户要求中途暂停确认，暂停后记录当前文件、已完成阶段和下一阶段，确认后从当前阶段继续，不要回到原始排产文件重跑。

补充约定：

- `format-main-sheet` 会在整理主表格式前再次刷新 `标台数` 公式和格式，避免新增计算列后表头、边框、数字格式不一致。
- `decompose-extra-summary` 生成的 `各线体分类明细表` 需要把 `订单数` 和 `标台数` 分开成独立表格展示；两个表格的“分类名称”顺序必须一致，`标台数` 四舍五入保留 1 位小数。
- `decompose-extra-summary` 还需要在 `排单分解表明细` 右侧“额外订单信息汇总”中追加 `产能规划` 来源指标：`外协烘道数量`、`滚筒喷粉数量`、`波轮喷粉数量`、`PCM板中需喷涂前门板的箱体数量`。其中喷粉指标从 `产能预算-喷涂` 表按项目/类别定位；PCM 箱体数量为 `改PCM箱体(未改门板)` 与 `改PCM箱体(未改门板)-金属粉` 之和。
- 最新结果参考：`W2627-W2631周排产计划明细(07月)628(3)_定编结果_20260630_153000_分类结果_20260630_205106.xlsx` 已在“排单分解表明细”中生成这 4 个产能规划指标；若后续文件缺少任一指标，先补跑 `decompose-extra-summary`，再进入 `classify`。
- `classify` 后必须检查“分类结果汇总表”是否存在，并统计“未分类数据”行数。未分类行数大于 0 时，返回结果文件并说明异常，询问用户是否需要继续补分类规则。

如果最新文件已经有“排单分解表明细”和“各线体分类明细表”，但没有“分类结果汇总表”，并且基础数据三项待补都为 0，则直接从 `classify` 继续。若基础数据仍有待补，先补基础数据，再重新执行受影响的标准台数、格式、分解、额外汇总和 `classify`。

## 6. 异常处理

- 文件不存在：报告绝对路径并暂停。
- 文件名以 `~$` 开头：要求关闭 Excel 并提供原文件。
- `outputs` 中存在对应的 `~$` 临时锁文件或打开工作簿失败：要求保存并关闭 Excel，再重试当前阶段；不要处理临时锁文件。
- 扩展名不支持：要求 `.xlsx` 或 `.xlsm` 主工作簿。
- 找不到主工作表：报告实际工作表名称并暂停。
- 找到多个主工作表：请用户选择。
- 查询表字段不符合要求：原样报告缺少字段，不尝试模糊映射。
- Excel 文件被占用：要求保存并关闭后重试当前阶段。
- 任一阶段失败：不要跳过，也不要从头重跑；保留最后成功输出并从失败阶段重试。
