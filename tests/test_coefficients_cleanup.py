import unittest

from openpyxl import Workbook

from dingbian_classifier.coefficients import cleanup_order_rows
from dingbian_classifier.logger import ProcessingLogger


class CleanupOrderRowsTest(unittest.TestCase):
    def test_deletes_blank_order_blank_line_and_excluded_material_rows(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["订单数", "物料编码", "线体"])
        sheet.append([100, "KEEP-1", "A线"])
        sheet.append([None, "BLANK-ORDER", "A线"])
        sheet.append([200, "BLANK-LINE", None])
        sheet.append([300, "Z4U6010100", "C线"])
        sheet.append([400, "KEEP-2", "H线"])

        result = cleanup_order_rows(sheet, ProcessingLogger())

        self.assertEqual(result.deleted_blank_order_rows, 1)
        self.assertEqual(result.deleted_blank_line_rows, 1)
        self.assertEqual(result.deleted_excluded_material_rows, 1)
        self.assertEqual(
            [sheet.cell(row_index, 2).value for row_index in range(2, sheet.max_row + 1)],
            ["KEEP-1", "KEEP-2"],
        )


if __name__ == "__main__":
    unittest.main()
