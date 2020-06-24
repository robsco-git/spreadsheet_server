import unittest
from .context import SpreadsheetServer, SpreadsheetClient
from time import sleep
import os
import shutil
import sys
import logging

EXAMPLE_SPREADSHEET = "example.ods"
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
TESTS_PATH = "./tests"
SHEET_NAME = "Sheet1"


class TestClient(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Copy the example spreadsheet from the tests directory into the spreadsheets
        # directory

        shutil.copyfile(
            TESTS_PATH + "/" + EXAMPLE_SPREADSHEET,
            SPREADSHEETS_PATH + "/" + EXAMPLE_SPREADSHEET,
        )

        cls.server = SpreadsheetServer(log_level=logging.CRITICAL)
        cls.server.run()

    @classmethod
    def tearDownClass(cls):
        cls.server.stop()
        os.remove(SPREADSHEETS_PATH + "/" + EXAMPLE_SPREADSHEET)

    def setUp(self):
        self.sc = SpreadsheetClient(EXAMPLE_SPREADSHEET)

    def tearDown(self):
        self.sc.disconnect()

    def test_connect_invalid_spreadsheet(self):
        try:
            SpreadsheetClient(EXAMPLE_SPREADSHEET + "z")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(
                str(e), "The requested spreadsheet was not found."
            )

            # Give the ThreadedTCPServer some time to shut down correctly before
            # the next test
            sleep(1)

    def test_get_sheet_names(self):
        sheet_names = self.sc.get_sheet_names()
        self.assertEqual(sheet_names, ["Sheet1"])

    def test_set_cell(self):
        self.sc.set_cells(SHEET_NAME, "A1", 5)
        a1 = self.sc.get_cells(SHEET_NAME, "A1")
        self.assertEqual(a1, 5)

    def test_set_cell_invalid_sheet(self):
        try:
            self.sc.set_cells(SHEET_NAME + "z", "A1", 5)
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Sheet name is invalid.")

    def test_get_cell(self):
        cell_value = self.sc.get_cells(SHEET_NAME, "C3")
        self.assertEqual(cell_value, 6)

    def test_get_cell_invalid_sheet(self):
        try:
            self.sc.get_cells(SHEET_NAME + "z", "C3")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Sheet name is invalid.")

    def test_get_invalid_cell_numeric(self):
        try:
            self.sc.get_cells(SHEET_NAME, 1)
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_missing_alpha(self):
        try:
            self.sc.get_cells(SHEET_NAME, "1")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_missing_numeric(self):
        try:
            self.sc.get_cells(SHEET_NAME, "A")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_missing_start_numeric(self):
        try:
            self.sc.get_cells(SHEET_NAME, "A:B2")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_missing_end_numeric(self):
        try:
            self.sc.get_cells(SHEET_NAME, "A1:B")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_missing_start_alpha(self):
        try:
            self.sc.get_cells(SHEET_NAME, "1:B2")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_missing_end_alpha(self):
        try:
            self.sc.get_cells(SHEET_NAME, "A1:2")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_negative(self):
        try:
            self.sc.get_cells(SHEET_NAME, "A-1:B2")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_numeric_too_large(self):
        try:
            self.sc.get_cells(SHEET_NAME, "A1048577")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_get_invalid_cell_alpha_too_large(self):
        try:
            self.sc.get_cells(SHEET_NAME, "AMK1")
            self.assertTrue(False)
        except RuntimeError as e:
            self.assertEqual(str(e), "Cell range is invalid.")

    def test_set_cell_row(self):
        cell_values = [4, 5, 6]
        self.sc.set_cells(SHEET_NAME, "A1:A3", cell_values)

        saved_values = self.sc.get_cells(SHEET_NAME, "A1:A3")
        self.assertEqual(cell_values, saved_values)

    def test_get_cell_column(self):
        cell_values = self.sc.get_cells(SHEET_NAME, "C1:C3")
        self.assertEqual(cell_values, [3, 3.5, 6])

    def test_get_cell_column_large_alpha(self):
        cell_values = self.sc.get_cells(SHEET_NAME, "AF5:AF186")
        self.assertEqual(cell_values, ["" for x in range(5, 186 + 1)])

    def test_set_cell_range(self):
        cell_values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        self.sc.set_cells(SHEET_NAME, "A1:C3", cell_values)

        saved_values = self.sc.get_cells(SHEET_NAME, "A1:C3")
        self.assertEqual(cell_values, saved_values)

    def test_save_spreadsheet(self):
        filename = "test.ods"
        self.sc.save_spreadsheet(filename)

        dir_path = os.path.dirname(os.path.realpath(__file__))

        saved_path = dir_path + "/../saved_spreadsheets/" + filename
        self.assertTrue(os.path.exists(saved_path))

        os.remove(saved_path)

    def test_unicode(self):
        for i in range(1000):
            self.sc.set_cells(SHEET_NAME, "A1", chr(i))
            cell = self.sc.get_cells(SHEET_NAME, "A1")
            try:
                int(chr(i))
                continue
            except ValueError:
                pass  # Not a number

            self.assertEqual(chr(i), cell)


if __name__ == "__main__":
    unittest.main()
