# 完整定编流程

## 命令约定

在 `E:\AI\dingbian` 中运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage <阶段>
```

每次从命令输出读取新的绝对文件路径，将它作为下一阶段的 `<当前文件>`。不要继续使用原始文件。

## 1. 系数

1. 运行 `prepare-coefficients`。
2. 检查输出。若 `coefficient_pending` 为 0，进入钣金阶段。
3. 若缺失且用户提供了唯一系数查询表，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage fill-coefficients --coefficient-lookup "<系数查询表>"
```

4. 若没有查询表，暂停并索取。
5. 查询后仍缺失，要求用户填写当前结果中的“系数补充”或“系数仍缺失”，保存并关闭 Excel。
6. 用户返回后运行 `apply-manual-coefficients`。仍缺失则继续暂停，不得进入下一阶段。

## 2. 钣金型号

1. 运行 `prepare-sheet-metal`。
2. 若缺失且用户提供了唯一钣金型号查询表，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage fill-sheet-metal --sheet-metal-lookup "<钣金型号查询表>"
```

3. 没有查询表时暂停索取。
4. 查询后仍缺失，要求用户填写“钣金型号补充”。
5. 用户返回后运行 `apply-manual-sheet-metal`。仍缺失则继续暂停。

## 3. 物料描述

1. 运行 `prepare-material-description`。
2. 若缺失且用户提供了唯一物料描述查询表，运行：

```powershell
python run_classifier.py --input "<当前文件>" --output-dir "outputs" --stage fill-material-description --material-description-lookup "<物料描述查询表>"
```

3. 没有查询表时暂停索取。
4. 查询后仍缺失，要求用户填写“物料描述补充”或“物料描述仍缺失”。
5. 用户返回后运行 `apply-manual-material-description`。仍缺失则继续暂停。

## 4. 标准台数与完整排单分解

基础数据全部补齐后，按以下固定顺序逐个运行；每个阶段必须使用上一阶段的新输出：

1. `prepare-standard-units`
2. `decompose-skd`
3. `decompose-rolling-remarks`
4. `decompose-t7p7t5p5-dreame`
5. `decompose-t9p9`
6. `decompose-t9p9-dryer`
7. `decompose-t10p10`
8. `decompose-c6-heat-pump-dryer`
9. `decompose-composite-penguin-c6`
10. `decompose-rolling-final`
11. `decompose-wave-basic`
12. `decompose-wave-final`
13. `decompose-extra-summary`
14. `classify`

不要把 `decompose-extra-summary` 提前到完整分解之前。

## 5. 异常处理

- 文件不存在：报告绝对路径并暂停。
- 文件名以 `~$` 开头：要求关闭 Excel 并提供原文件。
- 扩展名不支持：要求 `.xlsx` 或 `.xlsm` 主工作簿。
- 找不到主工作表：报告实际工作表名称并暂停。
- 找到多个主工作表：请用户选择。
- 查询表字段不符合要求：原样报告缺少字段，不尝试模糊映射。
- Excel 文件被占用：要求保存并关闭后重试当前阶段。
- 任一阶段失败：不要跳过，也不要从头重跑；保留最后成功输出并从失败阶段重试。
