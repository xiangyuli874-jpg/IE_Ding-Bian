import os
import unittest
from pathlib import Path
from unittest.mock import patch

from dingbian_classifier.excel_repair import resave_with_excel_if_available
from dingbian_classifier.logger import ProcessingLogger


class ExcelRepairTest(unittest.TestCase):
    def test_skips_intermediate_excel_resave_when_requested(self):
        with patch.dict(os.environ, {"DINGBIAN_SKIP_EXCEL_RESAVE": "1"}):
            with patch("dingbian_classifier.excel_repair.subprocess.run") as run:
                resave_with_excel_if_available(Path("C:/temporary/result.xlsx"), ProcessingLogger())

        run.assert_not_called()


if __name__ == "__main__":
    unittest.main()
