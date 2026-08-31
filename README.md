# 洗衣机厂效率预算与用人定编自动化工具

本项目用于处理《周排产计划明细》Excel 工作簿，自动识别 `W数字-W数字周排产明细` 主表，并围绕效率预算、用人定编和排产分类完成数据清理、系数补充、标台数计算、规则分类和结果汇总。

## 功能概览

- 自动查找周排产明细主工作表。
- 生成系数、钣金型号、物料描述等待补充工作表，支持查询表回填和人工回填。
- 支持一次性准备基础数据异常，连续摘取系数、钣金型号和物料描述待补项，避免前一项缺失时漏查后续项目。
- 默认保留订单数为空的预排未上单行；仅在明确要求时才清理该类订单。基础清理会删除线体为空及 `Z4U6010100`、`Z4U6010108`、`Z4U60501080` 三种排除编码，并刷新系数/钣金型号查找公式的当前行引用。
- 支持从钣金型号查询表或 BOM 表补齐钣金型号；BOM 表可先写入候选值等待人工确认，也可按需直接回填。
- 支持单独新增/刷新“标准单位/标准台数”相关列，方便在补齐基础数据后再计算标台数。
- 支持 SKD、滚筒备注、T7/P7/T5/P5/追觅、T9/P9、T10/P10、C6、复式、企鹅、滚筒收尾、波轮收尾和额外订单信息汇总等排单分解阶段。
- 根据主表已写入的 `类型` 生成分类结果汇总表和未分类数据，避免最终汇总再次套用旧规则造成已分解类型被覆盖。
- 可生成“排单分解表明细”“线体分类明细表”，汇总线体排单、标台数、转产时间和不良宽放工时，并对额外订单信息命中的钣金型号列上色。
- 额外订单信息汇总会追加产能规划来源指标：外协烘道数量、滚筒喷粉数量、波轮喷粉数量、PCM 板中需喷涂前门板的箱体数量。
- 支持对钣金型号为 `#N/A`、`N/A`、`0`、`0.0`、`0.00` 或 `-` 的行按物料编码定向回填物料描述，不覆盖钣金型号正常的订单描述。
- 支持全线体物料编码分类复核：`U605...` 识别干衣机、`U60101...` 识别单洗、`U60102...` 识别烘干；确定性错分自动纠正，无法唯一判断的冲突写入“物料编码分类复核”。
- 在处理结果中写入处理日志，并保留有限数量的历史备份。
- 清理旧结果时会跳过 Excel 打开的临时锁文件，避免误处理 `~$` 开头的临时文件。
- 提供 `电脑自动离线程序`，月底可双击启动分步向导，先用只读检查器确认主表和当前阶段，再按“自动查询优先、必要时手工补充”的方式跑完整月度流程。
- 提供 `手机自动离线程序`，可部署为 Streamlit 网页，手机、iPad 或电脑浏览器上传已补齐基础数据的 `.xlsx`/`.xlsm` 当前结果文件后下载处理结果。
- 提供 `skills/dingbian` Codex 技能，支持通过对话触发完整定编流程，并在每个补齐阶段后只读检查工作簿状态。

## 项目结构

```text
.
├── dingbian_classifier/      # 核心处理逻辑
├── skills/dingbian/          # Codex 定编自动化技能
├── 电脑自动离线程序/         # 电脑本地月度分步向导
├── 手机自动离线程序/         # Streamlit 网页处理入口
├── run_classifier.py         # 命令行入口
├── requirements.txt          # Python 依赖
└── README.md                 # 项目说明
```

`outputs/`、`inputs/` 中的 Excel 文件、各自动程序的 `logs/`、`tmp/`、`work/`、`__pycache__/`、Excel 业务数据、生成结果和向导续跑状态已通过 `.gitignore` 排除，不会默认提交到仓库。

## 环境准备

建议使用 Python 3.12 或兼容版本：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 基本用法

所有阶段都通过 `run_classifier.py` 执行，输入为周排产计划明细 Excel 文件，输出默认写入 `outputs` 目录。主排产工作簿只支持 `.xlsx` 或 `.xlsm`；`.xls` 可作为部分查询表或质量日报来源，不作为月度主文件入口。

```powershell
python run_classifier.py --input "<周排产计划明细.xlsx>" --output-dir "outputs" --stage classify
```

如果工作簿内存在多个候选主表，可以用 `--target-sheet` 明确指定要处理的 `W...周排产明细` 工作表。

## 常用处理阶段

一次性准备系数、钣金型号和物料描述基础数据异常。预排订单默认保留，只有线体为空或命中三种排除编码的订单会被清理：

```powershell
python run_classifier.py --input "<周排产计划明细.xlsx>" --output-dir "outputs" --stage prepare-foundation-data-preserve-order-blanks
```

仅当业务明确要求删除订单数为空的行时，清理订单数为空、线体为空和指定物料编码订单行，并刷新查找公式：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage cleanup-order-rows
```

生成“系数补充”并暂停：

```powershell
python run_classifier.py --input "<周排产计划明细.xlsx>" --output-dir "outputs" --stage prepare-coefficients
```

使用外部查询后的“系数查询表”回填主表系数：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage fill-coefficients --coefficient-lookup "<系数查询表.xlsx或xls>"
```

将手工维护在结果文件中的系数回填到主表：

```powershell
python run_classifier.py --input "<已手工补好系数的结果.xlsx>" --output-dir "outputs" --stage apply-manual-coefficients
```

保留订单数空白行、只准备三类基础异常：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage prepare-foundation-data-preserve-order-blanks
```

生成“钣金型号补充”：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage prepare-sheet-metal
```

使用外部查询后的“钣金型号查询表”回填主表：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage fill-sheet-metal --sheet-metal-lookup "<钣金型号查询表.xlsx或xls>"
```

使用 BOM 表把箱体组件候选写入“钣金型号补充”，等待人工确认：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage suggest-sheet-metal-bom --sheet-metal-bom-lookup "<钣金型号BOM表.xlsx>"
```

确认“钣金型号补充”后，将候选或手工值回填到主表：

```powershell
python run_classifier.py --input "<已确认钣金型号补充的结果.xlsx>" --output-dir "outputs" --stage apply-manual-sheet-metal
```

如已明确不需要人工确认，也可以用 BOM 表直接回填当前仍缺失或异常的钣金型号：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage fill-sheet-metal-bom --sheet-metal-bom-lookup "<钣金型号BOM表.xlsx>"
```

生成“物料描述补充”：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage prepare-material-description
```

使用外部查询后的“物料描述查询表”按“物料编码”回填主表：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage fill-material-description --material-description-lookup "<物料描述查询表.xlsx或txt/csv/tsv/xls导出>"
```

只对钣金型号异常的订单按物料编码补物料描述：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage fill-material-description-for-invalid-sheet-metal --material-description-lookup "<物料描述查询表.xlsx或txt/csv/tsv/xls导出>"
```

将手工维护在结果文件中的物料描述回填到主表：

```powershell
python run_classifier.py --input "<已手工补好物料描述的结果.xlsx>" --output-dir "outputs" --stage apply-manual-material-description
```

清理“物料描述补充”中确认无需上单的缺失订单行：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage cleanup-missing-material-description-rows
```

清理“系数补充”中确认无需处理的对应订单行：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage cleanup-coefficient-supplement-rows
```

刷新系数 VLOOKUP 当前行引用：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage refresh-coefficient-formulas
```

将仍缺系数、且需暂不参与定编分解的订单标记为跳过：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage skip-unresolved-coefficients
```

单独新增或刷新“标准单位/标准台数”相关列：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage prepare-standard-units
```

只调整辅助工作表顺序：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage reorder-sheets
```

执行完整分类流程：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage classify
```

生成额外订单信息汇总和线体分类明细：

```powershell
python run_classifier.py --input "<已完成排单分解的当前结果.xlsx>" --output-dir "outputs" --stage decompose-extra-summary
```

按物料编码复核单洗、烘干和干衣机大方向：

```powershell
python run_classifier.py --input "<已完成排单分解的当前结果.xlsx>" --output-dir "outputs" --stage audit-material-code-types
```

分类复核会修改主表 `类型` 时，必须重新生成线体分类明细表，再执行最终汇总：

```powershell
python run_classifier.py --input "<复核后的当前结果.xlsx>" --output-dir "outputs" --stage refresh-line-classification-detail
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage classify
```

## 推荐排单分解顺序

排单分解阶段会在主表中标记“类型”并刷新“排单分解表明细”。建议按下面顺序逐步处理，并将每一步输出作为下一步输入：

```powershell
python run_classifier.py --input "<当前结果.xlsx>" --output-dir "outputs" --stage decompose-skd
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-rolling-remarks
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-t7p7t5p5-dreame
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-t9p9
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-t9p9-dryer
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-t10p10
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-c6-heat-pump-dryer
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-composite-penguin-c6
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-rolling-final
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-wave-basic
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-wave-final
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage decompose-extra-summary
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage audit-material-code-types
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage refresh-line-classification-detail
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage classify
```

`decompose-extra-summary` 依赖主表中的“类型”列，先在滚筒/波轮分解完成后生成。物料编码复核如修正了类型，必须紧接着执行 `refresh-line-classification-detail`，不能沿用复核前的线体统计。

## 额外订单信息汇总

`decompose-extra-summary` 会在“排单分解表明细”右侧写入额外订单信息汇总，并新增或刷新“线体分类明细表”。当前汇总口径包括：

- 锥形筒：滚筒线，钣金型号含“锥形筒”。
- 波轮特殊内筒-10kg和9升10内筒：波轮线，钣金型号含“10kg波轮”或“9升10”。
- 波轮特殊内筒-8升9内筒：波轮线，钣金型号含“8升9”。
- 波轮箱体-10kg：波轮线，钣金型号含“10kg”。
- 波轮箱体-彩板：波轮线，钣金型号含“PCM”，并剔除波轮 CKD。
- 外协烘道数量：从“产能规划”中查找“外协烘道”右侧数量。
- 滚筒喷粉数量：从“产能规划”的“产能预算-喷涂”表中汇总项目为“滚筒”、类别为“喷涂”的数量。
- 波轮喷粉数量：从“产能规划”的“产能预算-喷涂”表中汇总项目为“波轮”、类别为“喷涂”的数量。
- PCM板中需喷涂前门板的箱体数量：汇总“外发/改PCM”下“改PCM箱体(未改门板)”和“改PCM箱体(未改门板)-金属粉”的数量。

“线体分类明细表”会按线体班组输出 `月度排单数`、`线体排单数`、`线体标台数`、`订单数`、`转产时间`、`返修损失/影响工时` 和右侧“订单不良宽放工时预算表”。标台数优先读取工作簿中的“标台数”列，缺失时按系数乘订单数计算，展示时保留 1 位小数。

当前线体分类口径以“线体分类明细表”为准；旧“各线体分类明细表”只作为历史结果来源。若某个线体班组模板缺少本次实际出现的分类机型，处理时允许在该线体班组内增加对应分类行，并在交付说明中列出新增的线体班组和分类。

A 线内销和 D 线外销公斤段细分以“钣金型号”为主，“物料描述”只做复核参考；同一公斤段会对“钣金型号”和“渠道”列使用一致颜色。波轮排单分解中原“塑料内销”统一输出为“塑料机”，主表“类型”、“排单分解表明细”和“线体分类明细表”保持一致。

最新交付校验要求 A 线内销、D 线外销公斤段不仅汇总到“线体分类明细表”，还必须同步回写主表“类型”列，并给主表“钣金型号”“渠道”两列按公斤段着色。主表按细分类型汇总的订单数应与线体分类明细中的月度排单数一致。

质量日报用于填写右侧“订单不良宽放工时预算表”时，`.xls` 或已知会触发 Excel 安全拦截的日报优先用非交互解析方式读取 `过程合格率&top问题` 工作表，再按 `100% - 合格率` 写入不良率，避免反复触发受保护视图或安全弹窗。

最终结果默认保留主表和结果表公式，只做必要的 Excel 重算和另存；不为了消除“发现内容有问题”弹窗而批量冻结为固定值。若另存固定值兼容版，需要在文件名和交付说明中明确标注。最终工作簿还需检查视图状态，尤其“线体分类明细表”冻结窗格应在表头附近，避免停在 `B89` 等深行位置导致无法正常滚动。

## 物料编码分类复核

`audit-material-code-types` 用于防止单洗、烘干和干衣机类型混分。复核规则以物料编码中 `U` 后的产品族为准：

- `U605...`：干衣机。
- `U60101...`：单洗。
- `U60102...`：烘干。

该阶段只自动修正规则能唯一确认的冲突，例如 `U605...` 且物料描述、备注或钣金型号能确认 T10/DWD10 干衣机时写为 `T10/P10干衣机`，`U60102...` 但仍停在普通外销大类时写为 `普通烘干`。不能唯一判断的订单会写入“物料编码分类复核”，保留原 `类型` 等待人工确认。

钣金型号异常但物料描述不可靠时，先运行 `fill-material-description-for-invalid-sheet-metal`，只针对钣金型号为 `#N/A`、`N/A`、`0`、`0.0`、`0.00` 或 `-` 的行按物料编码回填描述；钣金型号正常的行不会被覆盖。随后重新执行受影响的分解阶段，再运行 `decompose-extra-summary`、`audit-material-code-types`、`refresh-line-classification-detail` 和 `classify`。

最终 `classify` 会直接按主表当前 `类型` 汇总，生成“分类结果汇总表”和“未分类数据”；`类型` 为空的行进入“未分类数据”，不会再通过分类规则配置二次猜测。

## 电脑自动离线程序

`电脑自动离线程序` 用于电脑本地按需启动，不使用 Windows 任务计划程序，也不会后台常驻运行。

```powershell
电脑自动离线程序\启动月度处理.bat
```

启动后选择本次要处理的月度排产主 Excel 文件，主文件只支持 `.xlsx` 或 `.xlsm`。程序会先调用 `skills/dingbian/scripts/inspect_workbook.py` 做只读检查，确认唯一的 `W...周排产明细` 主表和当前待处理阶段，再按下面顺序运行分步向导：

1. 执行 `prepare-foundation-data-preserve-order-blanks`，默认保留预排未上单行，完成基础清理并生成系数、钣金型号和物料描述待补充表。
2. 查询后仍缺失时，暂停并打开结果文件，手工填写后再次双击继续。
3. 生成钣金型号待补充表，必要时选择钣金型号查询表自动回填。
4. 查询后仍缺失时，暂停并等待手工补充后继续。
5. 生成物料描述待补充表，必要时选择物料描述查询表自动回填。
6. 查询后仍缺失时，暂停并等待手工补充后继续。
7. 三类基础数据全部无待补后，按检查器返回的 `next_action` 决定从哪里继续。
8. 刷新“标准单位/标准台数”相关列。
9. 按 `电脑自动离线程序/monthly_flow.json` 中的最终 `stages` 顺序继续执行。

当前最终默认流程为：

```json
{
  "stages": [
    "prepare-standard-units",
    "format-main-sheet",
    "decompose-skd",
    "decompose-rolling-remarks",
    "decompose-t7p7t5p5-dreame",
    "decompose-t9p9",
    "decompose-t9p9-dryer",
    "decompose-t10p10",
    "decompose-c6-heat-pump-dryer",
    "decompose-composite-penguin-c6",
    "decompose-rolling-final",
    "decompose-wave-basic",
    "decompose-wave-final",
    "decompose-extra-summary",
    "audit-material-code-types",
    "refresh-line-classification-detail",
    "classify"
  ],
  "output_dir": "outputs"
}
```

最终流程会先刷新标台数、整理主表格式，再按滚筒线和波轮线完整分解，最后依次生成额外订单信息汇总、执行物料编码分类复核、刷新线体分类明细和最终分类结果。若检查器判断当前文件已经完成部分下游阶段，电脑向导会从 `prepare-standard-units`、`format-main-sheet`、`decompose-extra-summary` 或 `classify` 等合适位置继续，避免无意义地回到最初阶段重跑。

电脑向导会在暂停时保存 `电脑自动离线程序/wizard_state.json`，再次启动时可选择继续上次流程；全部完成后会自动清除该状态文件。每次运行会在 `电脑自动离线程序/logs` 下写入日志，记录检查器的 `next_action`、三类待补数量和未分类数量；处理完成后会自动打开 `outputs` 文件夹。若最终仍有未分类数据，完成弹窗会提示行数，请查看结果文件中的“未分类数据”。

## 手机网页程序

`手机自动离线程序` 是 Streamlit 网页入口，适合部署到 Streamlit Community Cloud 等平台后，用手机、iPad 或电脑浏览器上传 Excel 并下载结果。

本地测试：

```powershell
streamlit run 手机自动离线程序/web_app.py
```

部署到 Streamlit Community Cloud 时，Main file path 填：

```text
手机自动离线程序/web_app.py
```

当前网页默认执行 `手机自动离线程序/monthly_flow.json` 中的最终流程：

```json
{
  "stages": [
    "prepare-standard-units",
    "format-main-sheet",
    "decompose-skd",
    "decompose-rolling-remarks",
    "decompose-t7p7t5p5-dreame",
    "decompose-t9p9",
    "decompose-t9p9-dryer",
    "decompose-t10p10",
    "decompose-c6-heat-pump-dryer",
    "decompose-composite-penguin-c6",
    "decompose-rolling-final",
    "decompose-wave-basic",
    "decompose-wave-final",
    "decompose-extra-summary",
    "audit-material-code-types",
    "refresh-line-classification-detail",
    "classify"
  ],
  "output_dir": "outputs"
}
```

网页入口需要上传基础数据已经补齐、可进入最终流程的 `.xlsx` 或 `.xlsm` 当前结果文件，不接收 `.xls` 主文件。程序会刷新标台数和主表格式，继续执行完整分解、额外订单信息汇总和最终分类。可通过 Streamlit Secrets 设置 `app_password`，或通过环境变量 `DINGBIAN_APP_PASSWORD` 设置访问密码。

## Codex 定编技能

`skills/dingbian` 将现有命令行阶段组织成可复用的 Codex 技能，不重复实现业务分类规则。用户提供月度或周排产计划 Excel 后，可通过 `$dingbian` 或“帮我处理这个月排产并定编”等自然语言触发。

技能会：

1. 使用只读脚本检查主工作表、待补数据和最终结果表状态。
2. 优先并行准备系数、钣金型号、物料描述三类基础数据异常，再按标准台数、主表格式、完整排单分解、额外订单汇总、全线体物料编码分类复核、线体明细刷新、A/D 线公斤段回写校验和最终分类的顺序执行。
3. 缺少查询表或仍需人工补充时暂停，并明确返回当前结果文件和下一步操作。
4. 每次使用上一阶段的新输出继续处理，不修改原始工作簿。
5. 只有基础数据无待补、排单分解表和线体分类明细表已生成、物料编码分类复核已完成、额外订单汇总包含 4 个产能规划指标、A/D 线公斤段已回写主表并完成视图检查，且 `classify` 已生成“分类结果汇总表”和“未分类数据”后，才视为定编完成。

只读检查脚本也可以单独运行：

```powershell
python skills/dingbian/scripts/inspect_workbook.py "<工作簿.xlsx>"
```

脚本以 JSON 输出主工作表、数据行数、三类待补数量、未分类行数、最终结果表生成状态和建议的下一步操作。完整技能流程见 `skills/dingbian/references/workflow.md`。

## 支持的 stage

- `prepare-coefficients`：生成系数补充表。
- `prepare-foundation-data`：一次性生成系数、钣金型号、物料描述补充表，并完成订单行清理和公式刷新。
- `prepare-foundation-data-preserve-order-blanks`：一次性生成三类基础补充表，但保留订单数空白行。
- `cleanup-order-rows`：删除订单数空白、线体空白和指定物料编码订单行，并刷新系数/钣金型号公式。
- `cleanup-missing-material-description-rows`：删除物料描述缺失的未上单订单行，并刷新相关公式。
- `cleanup-coefficient-supplement-rows`：删除“系数补充”对应的订单行，并刷新相关公式。
- `refresh-coefficient-formulas`：刷新系数 VLOOKUP 当前行引用。
- `fill-coefficients`：从系数查询表回填。
- `apply-manual-coefficients`：从人工填写列回填系数。
- `skip-unresolved-coefficients`：标记系数仍缺失但暂不参与定编分解的订单行。
- `prepare-sheet-metal`：生成钣金型号补充表。
- `fill-sheet-metal`：从钣金型号查询表回填。
- `suggest-sheet-metal-bom`：按 BOM 号把箱体组件候选写入“钣金型号补充”，等待人工确认。
- `fill-sheet-metal-bom`：按 BOM 号直接回填当前缺失或异常的钣金型号。
- `apply-manual-sheet-metal`：从人工填写列回填钣金型号。
- `prepare-material-description`：生成物料描述补充表。
- `fill-material-description`：从物料描述查询表按物料编码回填。
- `fill-material-description-for-invalid-sheet-metal`：只对钣金型号异常行按物料编码回填物料描述。
- `apply-manual-material-description`：从人工填写列回填物料描述。
- `prepare-standard-units`：新增或刷新标准单位/标准台数相关列。
- `reorder-sheets`：调整辅助工作表顺序。
- `format-main-sheet`：整理主表格式。
- `decompose-skd`：执行 SKD 排单分解。
- `decompose-rolling-remarks`：执行滚筒备注相关排单分解。
- `decompose-t7p7t5p5-dreame`：执行 T7/P7/T5/P5/追觅规则分解。
- `decompose-t9p9`：执行 T9/P9 规则分解。
- `decompose-t9p9-dryer`：执行 T9/P9 干衣机规则分解。
- `decompose-t10p10`：执行 T10/P10 洗衣机和干衣机规则分解。
- `decompose-c6-heat-pump-dryer`：执行 C6 热泵干衣机规则分解。
- `decompose-composite-penguin-c6`：执行复式、企鹅和 C6 相关规则分解。
- `decompose-rolling-final`：执行滚筒收尾规则，补齐普通烘干、普通内销和外销等分类，并统计剩余未分类差异。
- `decompose-wave-basic`：执行波轮基础规则，覆盖 CKD、LG、塑料机、P7/P9 和 SKD 等分类。
- `decompose-wave-final`：执行波轮收尾规则，补齐内销铁皮变频、外销普通变频、内销铁皮和外销铁皮等分类，并统计剩余未分类差异。
- `decompose-extra-summary`：生成额外订单信息汇总和线体分类明细表，并对命中的钣金型号列上色。
- `audit-material-code-types`：按物料编码产品族复核并纠正确定性分类错误，剩余冲突写入“物料编码分类复核”。
- `refresh-line-classification-detail`：按复核后的主表 `类型` 重新生成“线体分类明细表”。
- `classify`：按主表 `类型` 汇总并生成结果表。

## 文件保留与回退

- 原始 Excel 不会被修改，程序会复制一份到输出目录再处理。
- `outputs` 目录只保留当前最新结果文件。
- 每次继续处理前，程序会先将当前结果备份到 `outputs/history`。
- `outputs/history` 默认最多保留 3 个备份。
- 清理旧结果和历史备份时，程序只按明确文件路径逐个删除文件，不使用递归批量删除。

## 注意事项

- 运行前请确认输入工作簿未被 Excel 独占锁定。
- Windows 环境下，如果本机可调用 Microsoft Excel，程序会尝试重新另存结果文件以提升兼容性；失败时会保留 openpyxl 生成的结果。
- 当前仓库默认不提交 Excel 文件。如需共享样例或模板，请先脱敏，再调整 `.gitignore` 白名单。
