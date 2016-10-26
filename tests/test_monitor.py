import unittest
from .context import MonitorThread, SpreadsheetServer
from time import sleep
import os

TEST_SS = "example.ods"
TEST_SS_MOVED = "example_moved.ods"
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"

class TestMonitor(unittest.TestCase):
    
    def setUp(self):
        self.spreadsheet_server = SpreadsheetServer()
        self.spreadsheet_server._SpreadsheetServer__start_soffice()
        self.spreadsheet_server._SpreadsheetServer__connect_to_soffice()
        self.spreadsheet_server._SpreadsheetServer__start_monitor_thread()
        self.monitor_thread = self.spreadsheet_server.monitor_thread
        while not self.monitor_thread.initial_scan():
            sleep(1)


    def tearDown(self):
        self.spreadsheet_server._SpreadsheetServer__stop_monitor_thread()
        self.spreadsheet_server._SpreadsheetServer__kill_libreoffice()
        self.spreadsheet_server._SpreadsheetServer__close_logfile()


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
        
if __name__ == '__main__':
    unittest.main()
