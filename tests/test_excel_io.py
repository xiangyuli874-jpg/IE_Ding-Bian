import unittest

from openpyxl import Workbook

from dingbian_classifier.excel_io import reset_auto_filter


class ExcelIoTest(unittest.TestCase):
    def test_reset_auto_filter_keeps_dropdowns_but_clears_active_criteria(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["线体", "订单数"])
        sheet.append(["A线", 100])
        sheet.append(["C线", 200])
        sheet.auto_filter.ref = "A1:B20"
        sheet.auto_filter.add_filter_column(0, [""])

        reset_auto_filter(sheet)

        self.assertEqual(sheet.auto_filter.ref, sheet.dimensions)
        self.assertEqual(list(sheet.auto_filter.filterColumn), [])


if __name__ == "__main__":
    unittest.main()
