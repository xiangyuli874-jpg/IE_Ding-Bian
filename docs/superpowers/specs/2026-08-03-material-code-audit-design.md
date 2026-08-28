# Material-code classification audit and description backfill

## Goal

Add a repeatable final classification audit that prevents single-wash, drying, and dryer products from being mixed. Then enrich rows whose sheet-metal model is `#N/A` or zero from the supplied material-description lookup, and rerun only the affected downstream results.

## Data rules

The product family is determined from the material-code body after `U`:

- `605...`: dryer
- `60101...`: single-wash
- `60102...`: drying

The audit checks the final `类型` against that product family. It auto-corrects only when the existing description, remark, or sheet-metal model identifies an existing unique detailed type. Otherwise, it writes the row to a review sheet and leaves the source type unchanged.

For the currently identified conflicts, the deterministic corrections are:

- `U605...` + T10/DWD10 dryer evidence -> `T10/P10干衣机`
- `U60102...` with no more-specific dry type -> `普通烘干`

## Workflow placement

1. Apply the supplied material-description lookup by material code only to rows with sheet-metal `#N/A`, `N/A`, `0`, `0.0`, or `0.00`.
2. Re-run rolling and wave decomposition stages so description-dependent rules are evaluated again.
3. Run the material-code audit after decomposition and before `decompose-extra-summary` creates the final result sheets.
4. Rebuild the decomposition detail, line classification detail, and classification result tables.
5. Reapply the already approved July quality-rate values because rebuilding the line-classification detail recreates its budget panel.

## Safety and verification

- Do not alter rows with valid sheet-metal models during the lookup step.
- Preserve existing formulas and styles; only the intended material-description and final type cells may change.
- Add regression tests for the three material-code families and the two deterministic corrections.
- Verify every material-code/type conflict after rerun is either corrected or listed in the dedicated audit sheet.
- Verify lookup coverage, type totals, final sheets, quality budget rates, and the rendered budget/classification views.

## Scope

No deletion of source rows or lookup data. Result workbooks remain in the existing `E:\AI\dingbian\outputs` root.
