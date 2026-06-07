# 洗衣机厂效率预算与用人定编自动化工具

本项目用于处理《周排产计划明细》Excel 工作簿，自动识别 `W数字-W数字周排产明细` 主表，并围绕效率预算、用人定编和排产分类完成数据清理、系数补充、标台数计算、规则分类和结果汇总。

## 功能概览

- 自动查找周排产明细主工作表。
- 生成系数、钣金型号等待补充工作表，支持查询表回填和人工回填。
- 支持 SKD、滚筒备注、T7/P7/T5/P5/追觅、T9/P9 等排单分解阶段。
- 根据工作簿内的分类规则配置生成分类结果、未分类数据和分类汇总表。
- 在处理结果中写入处理日志，并保留有限数量的历史备份。

## 项目结构

```text
.
├── dingbian_classifier/      # 核心处理逻辑
├── run_classifier.py         # 命令行入口
├── requirements.txt          # Python 依赖
└── README.md                 # 项目说明
```

`outputs/`、`__pycache__/`、Excel 业务数据和生成结果已通过 `.gitignore` 排除，不会默认提交到仓库。

## 环境准备

建议使用 Python 3.12 或兼容版本：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 基本用法

所有阶段都通过 `run_classifier.py` 执行，输入为周排产计划明细 Excel 文件，输出默认写入 `outputs` 目录。

```powershell
python run_classifier.py --input "<周排产计划明细.xlsx>" --output-dir "outputs" --stage classify
```

## 常用处理阶段

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

生成“钣金型号补充”：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage prepare-sheet-metal
```

使用外部查询后的“钣金型号查询表”回填主表：

```powershell
python run_classifier.py --input "<上一步输出.xlsx>" --output-dir "outputs" --stage fill-sheet-metal --sheet-metal-lookup "<钣金型号查询表.xlsx或xls>"
```

只调整辅助工作表顺序：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage reorder-sheets
```

执行完整分类流程：

```powershell
python run_classifier.py --input "<当前最新结果.xlsx>" --output-dir "outputs" --stage classify
```

## 支持的 stage

- `prepare-coefficients`：生成系数补充表。
- `fill-coefficients`：从系数查询表回填。
- `apply-manual-coefficients`：从人工填写列回填系数。
- `prepare-sheet-metal`：生成钣金型号补充表。
- `fill-sheet-metal`：从钣金型号查询表回填。
- `apply-manual-sheet-metal`：从人工填写列回填钣金型号。
- `reorder-sheets`：调整辅助工作表顺序。
- `format-main-sheet`：整理主表格式。
- `decompose-skd`：执行 SKD 排单分解。
- `decompose-rolling-remarks`：执行滚筒备注相关排单分解。
- `decompose-t7p7t5p5-dreame`：执行 T7/P7/T5/P5/追觅规则分解。
- `decompose-t9p9`：执行 T9/P9 规则分解。
- `classify`：执行分类并生成结果表。

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
