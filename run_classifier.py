from __future__ import annotations

import argparse
import sys
from pathlib import Path

from dingbian_classifier.exceptions import ClassifierError
from dingbian_classifier.pipeline import run


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="周排产明细自动识别与分类工具")
    parser.add_argument("--input", required=True, help="输入周排产计划明细 Excel 文件路径")
    parser.add_argument("--output-dir", default="outputs", help="输出目录，默认 outputs")
    parser.add_argument(
        "--stage",
        default="classify",
        choices=[
            "prepare-coefficients",
            "fill-coefficients",
            "apply-manual-coefficients",
            "prepare-sheet-metal",
            "fill-sheet-metal",
            "apply-manual-sheet-metal",
            "reorder-sheets",
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
            "classify",
        ],
        help="执行阶段：prepare-coefficients 生成系数补充；fill-coefficients 查询表回填；apply-manual-coefficients 手工系数回填；prepare-sheet-metal 生成标台数和钣金型号补充；fill-sheet-metal 查询表回填钣金型号；apply-manual-sheet-metal 手工回填钣金型号；reorder-sheets 调整辅助表顺序；format-main-sheet 格式整理；decompose-skd 排单分解SKD规则；decompose-rolling-remarks 排单分解滚筒CKD/三星/双滚筒规则；decompose-t7p7t5p5-dreame 排单分解T7/P7/T5/P5/追觅规则；decompose-t9p9 排单分解T9/P9规则；decompose-t9p9-dryer 排单分解T9/P9干衣机规则；decompose-t10p10 排单分解T10/P10规则；decompose-c6-heat-pump-dryer 排单分解C6热泵干衣机规则；decompose-composite-penguin-c6 排单分解复式/企鹅/C6规则；decompose-rolling-final 排单分解滚筒收尾规则；decompose-wave-basic 排单分解波轮第一组规则；decompose-wave-final 排单分解波轮收尾规则；classify 执行原分类流程",
    )
    parser.add_argument("--coefficient-lookup", help="系数查询表路径，仅 fill-coefficients 阶段需要")
    parser.add_argument("--sheet-metal-lookup", help="钣金型号查询表路径，仅 fill-sheet-metal 阶段需要")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        output_path = run(
            Path(args.input),
            Path(args.output_dir),
            stage=args.stage,
            coefficient_lookup=Path(args.coefficient_lookup) if args.coefficient_lookup else None,
            sheet_metal_lookup=Path(args.sheet_metal_lookup) if args.sheet_metal_lookup else None,
        )
    except ClassifierError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"[ERROR] 程序执行失败：{exc}", file=sys.stderr)
        return 1

    print(f"输出文件：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
