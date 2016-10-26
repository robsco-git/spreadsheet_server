# Copyright (C) 2016 Robert Scott

# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

import threading
from os import listdir
from os.path import isfile, isdir, join, exists
import logging
from time import sleep

class MonitorThread(threading.Thread):
    """Monitors the spreadsheet directory for changes."""

    def __init__(self, spreadsheets, locks, soffice, spreadsheets_path,
                 monitor_frequency):
        
        self._stop_thread = threading.Event()

        self.spreadsheets = spreadsheets
        self.locks = locks
        self.soffice = soffice
        self.spreadsheets_path = spreadsheets_path
        self.monitor_frequency = monitor_frequency

        self.done_scan = False # Done an initial scan or not

        super().__init__()
     

    def stop_thread(self):
        self._stop_thread.set()

        
    def stopped(self):
        return self._stop_thread.isSet()


    def initial_scan(self):
        return self.done_scan
    
        
    def __load_spreadsheet(self, doc):
        logging.info("Loading " + doc)
        self.spreadsheets[doc] = self.soffice.open_spreadsheet(
            self.spreadsheets_path + "/" + doc)
        self.locks[doc] = threading.Lock()

        
    def __unload_spreadsheet(self, doc):
        logging.info("Removing " + doc)
        self.locks[doc].acquire()
        self.spreadsheets[doc].close()
        self.spreadsheets.pop(doc, None)
        self.locks.pop(doc, None)

        
    def __check_added(self):
        """Check for new spreadsheets and loads them into LibreOffice."""

        for doc in self.docs:
            if doc[0] != '.': # Ignore hidden files
                found = False

                for key, value in self.spreadsheets.items():
                    if doc == key:
                        found = True
                        break

                if found == False:
                    self.__load_spreadsheet(doc)

                    
    def __check_removed(self):
        """Check for any deleted or removed spreadsheets and remove them from 
        LibreOffice.
        """
            
        removed_spreadsheets = []
        for key, value in self.spreadsheets.items():
            removed = True
            for doc in self.docs:
                if key == doc:
                    removed = False
                    break
            if removed:
                removed_spreadsheets.append(key)

        for doc in removed_spreadsheets:
            self.__unload_spreadsheet(doc)


    def __scan_directory(self, d):
        """Recursively scan a directory for spreadsheets."""

        dir_contents = listdir(d)

        for f in dir_contents:

            # Ignore particular files
            if f[:7] == ".~lock." or f == ".gitignore":
                continue

            full_path = join(d, f)
            if isfile(full_path):

                # Remove self.spreadsheets_path from the path
                relative_path = full_path.split(
                    self.spreadsheets_path)[1][1:]

                self.docs.append(relative_path)
            elif isdir(full_path):
                self.__scan_directory(full_path)

                
    def run(self):
        while not self.stopped():
            self.docs = []

            self.__scan_directory(self.spreadsheets_path)

            self.__check_removed()
            self.__check_added()

            self.done_scan = True

            sleep(self.monitor_frequency)
