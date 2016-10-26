import unittest
from .context import SpreadsheetServer, SpreadsheetClient
from time import sleep
import os

TEST_SS = "example.ods"
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
SHEET_NAME = "Sheet1"

class TestMonitor(unittest.TestCase):
    
    def setUp(self):
        self.server = SpreadsheetServer()
        self.server.run()

        self.sc = SpreadsheetClient(TEST_SS)


    def tearDown(self):
        self.sc.disconnect()
        self.server.stop()
        

    def test_get_sheet_names(self):
        sheet_names = self.sc.get_sheet_names()
        self.assertEqual(sheet_names, ["Sheet1"])


    def test_set_cell(self):
        self.sc.set_cells(SHEET_NAME, "A1", 5)
        a1 = self.sc.get_cells(SHEET_NAME, "A1")
        self.assertEqual(a1, 5)


    def test_get_cell(self):
        cell_value = self.sc.get_cells(SHEET_NAME, "C3")
        self.assertEqual(cell_value, 6)


    def test_set_cell_row(self):
        cell_values = [4, 5, 6]
        self.sc.set_cells(SHEET_NAME, "A1:A3", cell_values)

        saved_values = self.sc.get_cells(SHEET_NAME, "A1:A3")
        self.assertEqual(cell_values, saved_values)


    def test_get_cell_column(self):
        cell_values = self.sc.get_cells(SHEET_NAME, "C1:C3")
        self.assertEqual(cell_values, [3, 3.5, 6])


    def test_set_cell_range(self):
        cell_values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        self.sc.set_cells(SHEET_NAME, "A1:C3", cell_values)

        saved_values = self.sc.get_cells(SHEET_NAME, "A1:C3")
        self.assertEqual(cell_values, saved_values)


    def test_save_spreadsheet(self):
        filename = "test.ods"
        self.sc.save_spreadsheet(filename)

        dir_path = os.path.dirname(os.path.realpath(__file__))

        saved_path = dir_path + '/../saved_spreadsheets/' + filename
        self.assertTrue(os.path.exists(saved_path))
        

if __name__ == '__main__':
    unittest.main()
