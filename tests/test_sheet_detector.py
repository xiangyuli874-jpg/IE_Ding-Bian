import unittest

from dingbian_classifier.logger import ProcessingLogger
from dingbian_classifier.sheet_detector import find_target_sheet


class SheetDetectorTest(unittest.TestCase):
    def test_uses_explicit_target_sheet_when_provided(self):
        result = find_target_sheet(
            ["不加单周排产明细（不含2.4万）", "加单周排产明细（含2.4万）"],
            ProcessingLogger(),
            target_sheet_name="加单周排产明细（含2.4万）",
        )

        self.assertEqual(result, "加单周排产明细（含2.4万）")


if __name__ == "__main__":
    unittest.main()
