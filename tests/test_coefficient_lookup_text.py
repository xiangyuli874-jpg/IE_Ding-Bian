import tempfile
import unittest
from pathlib import Path

from dingbian_classifier.coefficients import read_lookup_rows


class TextCoefficientLookupTest(unittest.TestCase):
    def test_reads_sap_export_when_headers_follow_report_title(self):
        content = "\n".join(
            [
                "2026.07.25 动态清单显示 1",
                "",
                "\t工厂\t物料编码\t物料描述\t等级A\t等级B\t等级C\t等级D",
                "",
                "\t2000\tZ4U60101080848\t型号A\t2.384\t2.384\t2.384\t2.384",
                "\t2000\tZ4U60501080119\t型号B\t7.348\t6.441\t5.533\t4.625",
            ]
        )
        with tempfile.TemporaryDirectory() as directory:
            lookup = Path(directory) / "系数.xls"
            lookup.write_text(content, encoding="utf-16")

            rows = read_lookup_rows(lookup)

        self.assertEqual(
            rows,
            [("Z4U60101080848", "2.384"), ("Z4U60501080119", "4.625")],
        )


if __name__ == "__main__":
    unittest.main()
