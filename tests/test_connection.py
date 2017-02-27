import unittest
import threading
import os, shutil
from .context import SpreadsheetConnection, SpreadsheetServer
from signal import SIGTERM

import sys
PY2 = sys.version_info[0] == 2
PY3 = sys.version_info[0] == 3
if PY2:
    string_type = unicode
elif PY3:
    string_type = str
else:
    raise RuntimeError("Python version not supported.")


EXAMPLE_SPREADSHEET = "example.ods"
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
TESTS_PATH = "./tests"

class TestConnection(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Copy the example spreadsheet from the tests directory into the spreadsheets 
        # directory

        shutil.copyfile(
            TESTS_PATH + '/' + EXAMPLE_SPREADSHEET, 
            SPREADSHEETS_PATH + '/' + EXAMPLE_SPREADSHEET
        )

        cls.spreadsheet_server = SpreadsheetServer()
        cls.spreadsheet_server._SpreadsheetServer__start_soffice()
        cls.spreadsheet_server._SpreadsheetServer__connect_to_soffice()


    @classmethod
    def tearDownClass(cls):
        cls.spreadsheet_server._SpreadsheetServer__kill_libreoffice()
        cls.spreadsheet_server._SpreadsheetServer__close_logfile()
        os.remove(SPREADSHEETS_PATH + '/' + EXAMPLE_SPREADSHEET)
    
    
    def setUp(self):
        soffice = self.spreadsheet_server.soffice
        self.spreadsheet = soffice.open_spreadsheet(
            SPREADSHEETS_PATH + "/" + EXAMPLE_SPREADSHEET)

        lock = threading.Lock()
        self.ss_con = SpreadsheetConnection(self.spreadsheet, lock,
                                            self.spreadsheet_server.save_path)


    def tearDown(self):
        self.spreadsheet.close()
        
        
    def test_lock_spreadsheet(self):
        self.ss_con.lock_spreadsheet()
        self.assertTrue(self.ss_con.lock.locked())
        self.ss_con.unlock_spreadsheet()

        
    def test_unlock_spreadsheet(self):
        self.ss_con.lock_spreadsheet()
        status = self.ss_con.unlock_spreadsheet()
        self.assertTrue(status)
        self.assertFalse(self.ss_con.lock.locked())


    def test_unlock_spreadsheet_runtime_error(self):
        status = self.ss_con.unlock_spreadsheet()
        self.assertFalse(status)
        self.assertFalse(self.ss_con.lock.locked())
        

    def test_get_xy_index_first_cell(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"A1")
        self.assertEqual(alpha_index, 0)
        self.assertEqual(num_index, 0)


    def test_get_xy_index_Z26(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"Z26")
        self.assertEqual(alpha_index, 25)
        self.assertEqual(num_index, 25)


    def test_get_xy_index_aa3492(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"AA3492")
        self.assertEqual(alpha_index, 26)
        self.assertEqual(num_index, 3491)


    def test_get_xy_index_aaa1024(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"AAA1024")
        self.assertEqual(alpha_index, 702)
        self.assertEqual(num_index, 1023)


    def test_get_xy_index_aab739(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"AAB739")
        self.assertEqual(alpha_index, 703)
        self.assertEqual(num_index, 738)


    def test_get_xy_index_aba1(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"ABA1")
        self.assertEqual(alpha_index, 728)
        self.assertEqual(num_index, 0)


    def test_get_xy_index_abc123(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"ABC123")
        self.assertEqual(alpha_index, 730)
        self.assertEqual(num_index, 122)


    def test_get_xy_index_last_cell(self):
        alpha_index, num_index = self.ss_con._SpreadsheetConnection__get_xy_index(u"AMJ1048576")
        self.assertEqual(alpha_index, 1023)
        self.assertEqual(num_index, 1048575)


    def test_is_single_cell(self):
        status = self.ss_con._SpreadsheetConnection__is_single_cell(u"AMJ1")
        self.assertTrue(status)


    def test_is_not_single_cell(self):
        status = self.ss_con._SpreadsheetConnection__is_single_cell(u"A1:Z26")
        self.assertFalse(status)


    def test_check_not_single_cell(self):
        status = True
        try:
            self.ss_con._SpreadsheetConnection__check_single_cell(u"DD56:Z98")
        except ValueError:
            status = False
            
        self.assertFalse(status)
        

    def test_cell_to_index(self):
        d = self.ss_con._SpreadsheetConnection__cell_to_index(u"ABC945")

        self.assertTrue(d['row_index'] == 944)
        self.assertTrue(d['column_index'] == 730)


    def test_cell_range_to_index(self):
        d = self.ss_con._SpreadsheetConnection__cell_range_to_index(u"C9:Z26")

        self.assertTrue(d['row_start'] == 8)
        self.assertTrue(d['row_end'] == 25)
        
        self.assertTrue(d['column_start'] == 2)
        self.assertTrue(d['column_end'] == 25)


    def test_check_for_lock(self):
        status = False
        try:
            self.ss_con._SpreadsheetConnection__check_for_lock()
        except RuntimeError:
            status = True

        self.assertTrue(status)


    def test_check_numeric(self):
        value = self.ss_con._SpreadsheetConnection__convert_to_float_if_numeric("1")
        self.assertTrue(type(value) is float)


    def test_check_not_numeric(self):
        value = self.ss_con._SpreadsheetConnection__convert_to_float_if_numeric(
            u"123A")

        self.assertTrue(type(value) is str)


    def test_check_not_list(self):
        status = False
        try:
            self.ss_con._SpreadsheetConnection__check_list(1)
        except ValueError:
            status = True

        self.assertTrue(status)


    def test_check_1D_list(self):
        data = ["1","2","3"]
        data = self.ss_con._SpreadsheetConnection__check_1D_list(data)

        self.assertEqual(data, [1.0, 2.0, 3.0])

    def test_check_1D_list_when_2D(self):
        status = False
        data = [["1","2","3"], ["1", "2", "3"]]
        try:
            data = self.ss_con._SpreadsheetConnection__check_1D_list(data)
        except ValueError:
            status = True

        self.assertTrue(status)


    def test_invalid_sheet_name(self):
        self.ss_con.lock_spreadsheet()
        status = False
        try:
            self.ss_con.set_cells(u"1Sheet1", u"A1", 1)
        except ValueError:
            status = True
        self.ss_con.unlock_spreadsheet()

        self.assertTrue(status)
        

    def test_set_single_cell(self):
        self.ss_con.lock_spreadsheet()
        self.ss_con.set_cells(u"Sheet1", u"A1", 1)
        self.assertEqual(self.ss_con.get_cells(u"Sheet1", u"A1"), 1)
        self.ss_con.unlock_spreadsheet()


    def test_set_single_cell_list_of_data(self):
        self.ss_con.lock_spreadsheet()
        status = False
        try:
            self.ss_con.set_cells(u"Sheet1", u"A1", [9,1])
        except ValueError:
            status = True

        self.assertTrue(status)
        self.assertNotEqual(self.ss_con.get_cells(u"Sheet1", u"A1"), 9)
        self.ss_con.unlock_spreadsheet()


    def test_set_cell_range_columnn(self):
        self.ss_con.lock_spreadsheet()
        self.ss_con.set_cells(u"Sheet1", u"A1:A5", [1, 2, 3, 4, 5])
        self.assertEqual(self.ss_con.get_cells(u"Sheet1", u"A1:A5"),
                         (1.0, 2.0, 3.0, 4.0, 5.0))
        self.ss_con.unlock_spreadsheet()


    def test_set_cell_range_row(self):
        self.ss_con.lock_spreadsheet()
        self.ss_con.set_cells(u"Sheet1", u"A1:E1", [1, 2, 3, 4, 5])
        self.assertEqual(self.ss_con.get_cells(u"Sheet1", u"A1:E1"),
                         (1.0, 2.0, 3.0, 4.0, 5.0))
        self.ss_con.unlock_spreadsheet()
        
        
    def test_set_cell_range_2D(self):
        self.ss_con.lock_spreadsheet()
        self.ss_con.set_cells(u"Sheet1", u"A1:B2", [[1,2],[3,4]])
        self.assertEqual(self.ss_con.get_cells(u"Sheet1", u"A1:B2"),
                         ((1.0,2.0),(3.0,4.0)))
        self.ss_con.unlock_spreadsheet()



    def test_set_cell_range_2D_incorrect_data(self):
        self.ss_con.lock_spreadsheet()
        status = False
        try:
            self.ss_con.set_cells(u"Sheet1", u"A1:B2", [9, 9, 9, 9])
        except ValueError:
            status = True

        self.assertTrue(status)
        
        self.assertNotEqual(self.ss_con.get_cells(u"Sheet1", u"A1:B2"),
                         ((9.0,9.0),(9.0,9.0)))
        self.ss_con.unlock_spreadsheet()


    def test_get_sheet_names(self):
        sheet_names = self.ss_con.get_sheet_names()
        self.assertEqual(sheet_names, [u'Sheet1']

    )

    def test_save_spreadsheet(self):
        path = "./saved_spreadsheets/" + EXAMPLE_SPREADSHEET + ".new"

        if os.path.exists(path):
            os.remove(path)
        
        self.ss_con.lock_spreadsheet()
        status = self.ss_con.save_spreadsheet(EXAMPLE_SPREADSHEET + ".new")
        self.assertTrue(status)
        self.assertTrue(os.path.exists(path))
        self.ss_con.unlock_spreadsheet()


    def test_save_spreadsheet_no_lock(self):
        path = "./saved_spreadsheets/" + EXAMPLE_SPREADSHEET + ".new"

        if os.path.exists(path):
            os.remove(path)

        status = self.ss_con.save_spreadsheet(EXAMPLE_SPREADSHEET + ".new")
        self.assertFalse(status)
        self.assertFalse(os.path.exists("./saved_spreadsheets/" + EXAMPLE_SPREADSHEET +
                                       ".new"))

        
if __name__ == '__main__':
    unittest.main()
