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
import hashlib

class MonitorThread(threading.Thread):
    """Monitors the spreadsheet directory for changes."""

    def __init__(self, spreadsheets, locks, hashes, soffice, spreadsheets_path,
                 monitor_frequency, reload_on_disk_change):
        
        self._stop_thread = threading.Event()

        self.spreadsheets = spreadsheets
        self.locks = locks
        self.hashes = hashes
        self.soffice = soffice
        self.spreadsheets_path = spreadsheets_path
        self.monitor_frequency = monitor_frequency
        self.reload_on_disk_change = reload_on_disk_change

        self.done_scan = False # Done an initial scan or not

        super().__init__()
     

    def stop_thread(self):
        self._stop_thread.set()

        
    def stopped(self):
        return self._stop_thread.isSet()


    def initial_scan(self):
        return self.done_scan


    def __get_full_path(self, doc):
        return join(self.spreadsheets_path, doc)

        
    def __load_spreadsheet(self, doc):
        logging.info("Loading " + doc["path"])

        self.spreadsheets[doc["path"]] = self.soffice.open_spreadsheet(
            self.__get_full_path(doc["path"])
        )
        self.locks[doc["path"]] = threading.Lock()
        self.hashes[doc["path"]] = doc["hash"]

        
    def __unload_spreadsheet(self, doc_path):
        logging.info("Removing " + doc_path)
        self.locks[doc_path].acquire()
        self.spreadsheets[doc_path].close()
        self.spreadsheets.pop(doc_path, None)
        self.locks.pop(doc_path, None)
        self.hashes.pop(doc_path, None)

        
    def __check_added(self):
        """Check for new spreadsheets and loads them into LibreOffice."""

        for doc in self.docs:
            if doc["path"][0] != '.': # Ignore hidden files
                load = True # Default to loading the spreadsheet

                for key, value in self.spreadsheets.items():
                    if doc["path"] == key:

                        # Check if the file has been modified
                        # Does the file now have a differnet hash?

                        if (self.reload_on_disk_change
                            and doc["hash"] != self.hashes[doc["path"]]):
                            
                            self.__unload_spreadsheet(doc["path"])
                        else:
                            load = False
                            
                        break

                if load:
                    self.__load_spreadsheet(doc)

                    
    def __check_removed(self):
        """Check for any deleted or removed spreadsheets and remove them from 
        LibreOffice.
        """
            
        removed_spreadsheets = []
        for key, value in self.spreadsheets.items():
            removed = True
            for doc in self.docs:
                if key == doc["path"]:
                    removed = False
                    break
            if removed:
                removed_spreadsheets.append(key)

        for doc_path in removed_spreadsheets:
            self.__unload_spreadsheet(doc_path)


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

                # Calculate the MD5 hash for the file
                hasher = hashlib.md5()
                with open(self.__get_full_path(relative_path), 'rb') as afile:
                    buf = afile.read()
                    hasher.update(buf)
                    h = hasher.hexdigest()
                
                self.docs.append({"path": relative_path, "hash": h})
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
