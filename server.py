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

import logging
import pyoo
from com.sun.star.connection import NoConnectException
import subprocess
import threading
import socketserver
from socket import SHUT_RDWR 
from os import listdir
from os.path import isfile, join
from time import sleep
from connection import SpreadsheetConnection
import json
import traceback
import textwrap

SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
MONITOR_THREAD_FREQ = 60 # In seconds

class ThreadedTCPRequestHandler(socketserver.BaseRequestHandler):
    def __send(self, msg):
        """Convert a message to JSON and send it to the client.

        The messages are sent as utf-8 encoded bytes
        """

        json_msg = json.dumps(msg)
        json_bytes = bytes(json_msg, "utf-8")

        self.request.send(json_bytes)
        
        logging.debug("Sent: " + json.dumps(msg))


    def __receive(self):
        """Receive a message from the client, decode it from JSON and return.
        
        The received messages are utf-8 encoded bytes. False is returned on 
        failure to connect to the client, otherwise a string of the message is
        returned.
        """
        
        recv = self.request.recv(4096)
        if recv == b'':
            # The connection is closed.
            return False

        recv_json = str(recv, encoding="utf-8")
        recv_string = json.loads(recv_json)
        
        logging.debug("Received: " + str(recv_string))
        return recv_string


    def __make_connection(self):
        """Handle first request to server and check that it adheres to the
        protocol.
        """
        
        data = self.__receive()
        if (data[0] != "SPREADSHEET"):
            raise RuntimeError("Received incorrect connection string.")

        self.con = SpreadsheetConnection(
            server.spreadsheets[data[1]], server.locks[data[1]])
        
        self.__send("OK")
        self.con.lock_spreadsheet()


    def __close_connection(self):
        """Unlock the spreadsheet and close the connection to the client."""
        
        try:
            if self.con.lock.locked:
                self.con.unlock_spreadsheet()
        except UnboundLocalError:
            # con was never created.
            pass
            
        try:
            self.request.shutdown(SHUT_RDWR)
        except OSError:
            # The client has already disconnected.
            pass

        self.request.close()


    def __main_loop(self):
        while True:
            data = self.__receive()
            
            if data == False:
                # The connection has been lost.
                break
            
            elif data[0] == "SET":
                self.con.set_cells(data[1], data[2], data[3])
                self.__send("OK")
                
            elif data[0] == "GET":
                cells = self.con.get_cells(data[1], data[2])

                if cells != False:
                    self.__send(cells)
                else:
                    self.__send("ERROR")
                    
            elif data[0] == "SAVE":
                self.con.save_spreadsheet(data[1])
                self.__send("OK")

    
    def handle(self):
        """Make a connection to the client, run the main protocol loop and
        close the connection.
        """

        self.__make_connection()
        self.__main_loop()
        self.__close_connection()

        
class MonitorThread(threading.Thread):
    """Monitors the spreadsheet directory for changes."""
    
    def __load_spreadsheet(self, doc):
        logging.info("Loading " + doc)
        server.spreadsheets[doc] = soffice.open_spreadsheet(
            SPREADSHEETS_PATH + "/" + doc)
        server.locks[doc] = threading.Lock()

        
    def __unload_spreadsheet(self, doc):
        logging.info("Removing " + doc)
        server.locks[doc].acquire()
        server.spreadsheets[doc].close()
        server.spreadsheets.pop(doc, None)
        server.locks.pop(doc, None)

        
    def __check_added(self):
        """Check for new spreadsheets and loads them into LibreOffice."""
                    
        for doc in self.docs:
            if doc[0] != '.': # Ignore hidden files
                found = False

                for key, value in server.spreadsheets.items():
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
        for key, value in server.spreadsheets.items():
            removed = True
            for doc in self.docs:
                if key == doc:
                    removed = False
                    break
            if removed:
                removed_spreadsheets.append(key)

        for doc in removed_spreadsheets:
            self.__unload_spreadsheet(doc)

    
    def run(self):
        while True:
            spreadsheets = listdir(SPREADSHEETS_PATH)
            self.docs = [f for f in spreadsheets if
                    isfile(join(SPREADSHEETS_PATH, f))]

            self.__check_removed()
            self.__check_added()
            
            sleep(MONITOR_THREAD_FREQ)


class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    pass


if __name__ == "__main__":

    # Set up logging.
    LOG_FILE = './log/server.log'
    SOFFICE_LOG = './log/soffice.log'
    logging.basicConfig(format='%(asctime)s:%(levelname)s:%(message)s',
                        datefmt='%Y%m%d %H:%M:%S',
                        filename=LOG_FILE,
                        level=logging.DEBUG)
    
    logging.info('Starting calc_server.')

    logging.info('Starting the soffice process.')
    command = '/usr/bin/soffice --accept="pipe,name=' + SOFFICE_PIPE +';urp;"\
    --norestore --nologo --nodefault --headless'
    
    logfile = open(SOFFICE_LOG, "w")
    subprocess.Popen(command, shell=True, stdout=logfile, stderr=logfile)

    # Make a connection to soffice and fail if it can not connect
    MAX_ATTEMPTS = 60
    attempt = 0
    while 1:
        if attempt > MAX_ATTEMPTS: # soffice process isin't coming up
            raise RuntimeError("Could not connect to soffice process.")

        try:
            soffice = pyoo.Desktop(pipe=SOFFICE_PIPE)
            logging.info("Connected to soffice.")
            break
        
        except NoConnectException:
            attempt += 1
            sleep(1)

    # Start server initialisation
    HOST, PORT = "localhost", 5555
    server = ThreadedTCPServer((HOST, PORT), ThreadedTCPRequestHandler)
    
    # create documents for each file in ./spreadsheets
    files = listdir(SPREADSHEETS_PATH)
    docs = [f for f in files if isfile(join(SPREADSHEETS_PATH, f))]

    server.spreadsheets = {}
    server.locks = {} # A lock for each spreadsheet
    for doc in docs:
        if doc[0] == '.' :
            continue
        
        logging.info("Loading " + doc)
        server.spreadsheets[doc] = soffice.open_spreadsheet(
            SPREADSHEETS_PATH + "/" + doc)
        server.locks[doc] = threading.Lock()

    # Start the main server thread. This server thread will start a
    # new thread to handle each client connection.
    server_thread = threading.Thread(target=server.serve_forever)
    server_thread.daemon = False # Gracefully stop child threads
    server_thread.start()
    logging.info("Server thread running. Waiting on connections...")

    # This thread monitors the SPREADSHEETS directory to add or remove
    # spreadsheets
    monitor_thread = MonitorThread()
    monitor_thread.daemon = True
    monitor_thread.start()

    print("""calc_server version 69, Copyright (C) year name of author
calc_server comes with ABSOLUTELY NO WARRANTY. This is free software, and you are welcome
to redistribute it under certain conditions; type `show c'
for details.""")
