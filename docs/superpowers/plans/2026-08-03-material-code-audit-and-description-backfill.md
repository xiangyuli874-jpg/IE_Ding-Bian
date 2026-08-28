# Material-code audit and description backfill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Prevent single-wash, drying, and dryer categories from mixing; enrich invalid-sheet-metal rows from the supplied material-description lookup; and regenerate affected results.

**Architecture:** Add a targeted description-backfill stage for invalid sheet-metal rows and a material-code audit stage after decomposition. Existing rules stay authoritative; the audit only makes verified deterministic corrections and writes remaining conflicts to a review sheet.

**Tech Stack:** Python 3.12, openpyxl pipeline, Microsoft Excel calculation, unittest.

## Global Constraints

- Write result workbooks only in the existing `E:\AI\dingbian\outputs` root; do not alter source files or delete rows.
- Product family is based on the material-code body after `U`: `605` dryer, `60101` single-wash, `60102` drying.
- Reapply July quality rates after rebuilding `线体分类明细表`.

---

### Task 1: Description lookup limited to invalid sheet metal

**Files:**

- Modify: `dingbian_classifier/material_description.py`
- Modify: `dingbian_classifier/pipeline.py`
- Modify: `run_classifier.py`
- Create: `tests/test_material_description_invalid_sheet_metal.py`

**Interfaces:**

- Produces `fill_material_descriptions_for_invalid_sheet_metal(workbook, values_workbook, target_sheet_name, lookup_path, logger) -> MaterialDescriptionFillResult`.
- Registers CLI stage `fill-material-description-for-invalid-sheet-metal`.

- [ ] **Step 1: Write the failing test**

```python
def test_fills_lookup_description_only_when_sheet_metal_is_invalid():
    rows = [
        ["Z4U60101000001", "#N/A", "old"],
        ["Z4U60101000002", "500滚筒", "keep"],
        ["Z4U60101000003", 0, ""],
    ]
    result = fill_material_descriptions_for_invalid_sheet_metal(...)
    assert result.applied_rows == 2
    assert descriptions == ["lookup-1", "keep", "lookup-3"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_material_description_invalid_sheet_metal -v`

Expected: FAIL because the stage does not exist.

- [ ] **Step 3: Write minimal implementation**

```python
for row_index in range(2, formula_sheet.max_row + 1):
    if _is_invalid_sheet_metal(values_sheet.cell(row_index, metal_col).value):
        description = lookup.get(normalize_material_code(...))
        if description:
            formula_sheet.cell(row_index, description_col).value = description
            values_sheet.cell(row_index, description_col).value = description
```

Register the stage using the same workbook save and Excel-repair path as `fill-material-description`.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests.test_material_description_invalid_sheet_metal -v`

Expected: PASS; the valid-model row remains unchanged.

- [ ] **Step 5: Commit**

```bash
git add dingbian_classifier/material_description.py dingbian_classifier/pipeline.py run_classifier.py tests/test_material_description_invalid_sheet_metal.py
git commit -m "feat: backfill descriptions for invalid sheet metal"
```

### Task 2: Material-code category audit

**Files:**

- Modify: `dingbian_classifier/decomposition.py`
- Modify: `dingbian_classifier/pipeline.py`
- Modify: `run_classifier.py`
- Create: `tests/test_material_code_type_audit.py`

**Interfaces:**

- Produces `audit_and_correct_material_code_types(workbook, values_workbook, target_sheet_name, logger) -> MaterialCodeAuditResult`.
- Registers CLI stage `audit-material-code-types` and produces `物料编码分类复核`.

- [ ] **Step 1: Write the failing test**

```python
def test_audit_corrects_verified_dryer_and_drying_conflicts():
    rows = [
        ["Z3U60501080000", "T10热泵干衣机", "追觅", "普通内销"],
        ["Z4U60102080139", "TWF120...", "委内瑞拉", "外销"],
    ]
    result = audit_and_correct_material_code_types(...)
    assert types == ["T10/P10干衣机", "普通烘干"]
    assert result.corrected_rows == 2
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m unittest tests.test_material_code_type_audit -v`

Expected: FAIL because the audit stage does not exist.

- [ ] **Step 3: Write minimal implementation**

```python
if kind == "dryer" and type_name == DOMESTIC_TYPE and (
    "T10" in sheet_metal.upper() or description.upper().startswith("DWD10")
):
    return T10P10_DRYER_TYPE
if kind == "dry" and type_name == EXPORT_TYPE:
    return ORDINARY_DRY_TYPE
```

Write unresolved conflicts to `物料编码分类复核` with row number, code, description, remark, line, type, family, and audit result.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m unittest tests.test_material_code_type_audit -v`

Expected: PASS; verified conflicts correct and unresolved conflicts remain reviewable.

- [ ] **Step 5: Commit**

```bash
git add dingbian_classifier/decomposition.py dingbian_classifier/pipeline.py run_classifier.py tests/test_material_code_type_audit.py
git commit -m "feat: audit material code categories"
```

### Task 3: Persist the workflow and regenerate the current result

**Files:**

- Modify: `skills/dingbian/SKILL.md`
- Modify: `skills/dingbian/references/workflow.md`

**Interfaces:**

- Consumes the two new stages and current quality-rate values.
- Produces refreshed descriptions, corrected types, audit sheet, rebuilt results, and preserved quality budget rates.

- [ ] **Step 1: Add workflow instructions**

Document the order: invalid-sheet-metal description fill -> all affected decomposition -> material-code audit -> `decompose-extra-summary` -> `classify` -> quality-rate reapply.

- [ ] **Step 2: Run current workbook stages**

```powershell
python run_classifier.py --input "<current>" --output-dir "E:\AI\dingbian\outputs" --stage fill-material-description-for-invalid-sheet-metal --material-description-lookup "E:\A_IE_xiangyu\（A）定编预算\26年\8月\物料描述补充表.xlsx"
python run_classifier.py --input "<output>" --output-dir "E:\AI\dingbian\outputs" --stage audit-material-code-types
python run_classifier.py --input "<output>" --output-dir "E:\AI\dingbian\outputs" --stage decompose-extra-summary
python run_classifier.py --input "<output>" --output-dir "E:\AI\dingbian\outputs" --stage classify
```

- [ ] **Step 3: Reapply July quality values**

Write the approved July rates to `线体分类明细表!M4:M11`, preserving the existing P/Q formulas.

- [ ] **Step 4: Verify**

```powershell
python -m unittest discover -s tests -v
python "E:\AI\dingbian\skills\dingbian\scripts\inspect_workbook.py" "<final-output>"
```

Confirm the audit sheet has no unresolved conflict without a review record, all final sheets exist, and visual render is legible.

- [ ] **Step 5: Commit**

```bash
git add skills/dingbian/SKILL.md skills/dingbian/references/workflow.md docs/superpowers/plans/2026-08-03-material-code-audit-and-description-backfill.md
git commit -m "docs: add material code audit workflow"
```
