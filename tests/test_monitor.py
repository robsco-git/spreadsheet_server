import unittest
from .context import MonitorThread, SpreadsheetServer, SpreadsheetClient
from time import sleep
import os
import shutil

TEST_SS = "example.ods"
TEST_SS_MOVED = "example_moved.ods"
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
SAVED_SPREADSHEETS_PATH = "./saved_spreadsheets"
SHEET_NAME = "Sheet1"

class TestMonitor(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.spreadsheet_server = SpreadsheetServer()
        cls.spreadsheet_server._SpreadsheetServer__start_soffice()
        cls.spreadsheet_server._SpreadsheetServer__connect_to_soffice()


    @classmethod
    def tearDownClass(cls):
        cls.spreadsheet_server._SpreadsheetServer__kill_libreoffice()
        cls.spreadsheet_server._SpreadsheetServer__close_logfile()
        
    
    def setUp(self):
        self.spreadsheet_server._SpreadsheetServer__start_monitor_thread()
        self.monitor_thread = self.spreadsheet_server.monitor_thread
        while not self.monitor_thread.initial_scan():
            sleep(0.5)


    def tearDown(self):
        self.spreadsheet_server._SpreadsheetServer__stop_monitor_thread()
        

    def test_unload_spreadsheet(self):
        self.monitor_thread._MonitorThread__unload_spreadsheet(TEST_SS)
        
        spreadsheets = [
            key for key, value in self.monitor_thread.spreadsheets.items()
        ]
        
        locks = [
            key for key, value in self.monitor_thread.locks.items()
        ]

        self.assertTrue(TEST_SS not in spreadsheets)
        self.assertTrue(TEST_SS not in locks)

        
    def test_check_added_already_exists(self):
        self.monitor_thread._MonitorThread__check_added()

        spreadsheets = [
            key for key, value in self.monitor_thread.spreadsheets.items()
        ]
        
        locks = [
            key for key, value in self.monitor_thread.locks.items()
        ]

        self.assertTrue(TEST_SS in spreadsheets)
        self.assertTrue(TEST_SS in locks)


    def test_check_removed_when_renamed(self):
        # Rename example.ods to example_moved.ods
        
        current_loc = SPREADSHEETS_PATH + '/' + TEST_SS
        moved_loc = SPREADSHEETS_PATH + '/' + TEST_SS_MOVED

        os.rename(current_loc, moved_loc)

        self.monitor_thread.docs = []
        self.monitor_thread._MonitorThread__scan_directory(SPREADSHEETS_PATH)
        self.monitor_thread._MonitorThread__check_removed()
        self.monitor_thread._MonitorThread__check_added()

        spreadsheets = [
            key for key, value in self.monitor_thread.spreadsheets.items()
        ]
        
        locks = [
            key for key, value in self.monitor_thread.locks.items()
        ]

        self.assertTrue(TEST_SS not in spreadsheets)
        self.assertTrue(TEST_SS not in locks)
        self.assertTrue(TEST_SS_MOVED in spreadsheets)
        self.assertTrue(TEST_SS_MOVED in locks)

        # Move it back to where it was
        os.rename(moved_loc, current_loc)


    def test_change_file_hash(self):
        # Save the example file with a modification

        current_loc = SPREADSHEETS_PATH + '/' + TEST_SS
        moved_loc = SPREADSHEETS_PATH + '/' + TEST_SS_MOVED
        shutil.copyfile(current_loc, moved_loc)

        self.spreadsheet_server._SpreadsheetServer__start_threaded_tcp_server()
        
        self.sc = SpreadsheetClient(TEST_SS_MOVED)
        self.sc.set_cells(SHEET_NAME, "A1", 5)
        self.sc.save_spreadsheet(TEST_SS_MOVED)
        self.sc.disconnect()

        current_loc = SAVED_SPREADSHEETS_PATH + '/' + TEST_SS_MOVED
        moved_loc = SPREADSHEETS_PATH + '/' + TEST_SS_MOVED

        hash_before = self.monitor_thread.hashes[TEST_SS_MOVED]

        os.rename(current_loc, moved_loc)

        # Run a scan manually
        self.monitor_thread.docs = []
        self.monitor_thread._MonitorThread__scan_directory(SPREADSHEETS_PATH)
        self.monitor_thread._MonitorThread__check_removed()
        self.monitor_thread._MonitorThread__check_added()

        hash_after = self.monitor_thread.hashes[TEST_SS_MOVED]
        self.assertNotEqual(hash_before, hash_after)
        
        self.sc = SpreadsheetClient(TEST_SS_MOVED)
        cell = self.sc.get_cells(SHEET_NAME, "A1")
        self.sc.save_spreadsheet(TEST_SS_MOVED)
        self.sc.disconnect()

        self.assertEqual(cell, 5)

        self.spreadsheet_server._SpreadsheetServer__stop_threaded_tcp_server()
        os.remove(moved_loc)
        
        
if __name__ == '__main__':
    unittest.main()
