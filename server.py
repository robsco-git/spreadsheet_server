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
import subprocess
import threading
from time import sleep
from request_handler import ThreadedTCPRequestHandler, ThreadedTCPServer
from monitor import MonitorThread
from signal import SIGTERM, pause

LOG_FILE = './log/server.log'
SOFFICE_LOG = './log/soffice.log'
SOFFICE_HOST, SOFFICE_PORT = "localhost", 5555
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
MONITOR_FREQ = 5 # In seconds


class SpreadsheetServer():

    def __init__(self, log_file=LOG_FILE, soffice_log=SOFFICE_LOG,
                 soffice_host=SOFFICE_HOST, soffice_port=SOFFICE_PORT,
                 soffice_pipe=SOFFICE_PIPE,
                 spreadsheets_path=SPREADSHEETS_PATH,
                 monitor_frequency=MONITOR_FREQ):
        
        self.log_file = log_file
        self.soffice_log = soffice_log
        self.soffice_host = soffice_host
        self.soffice_port = soffice_port
        self.soffice_pipe = soffice_pipe
        self.monitor_frequency = monitor_frequency
        
        self.spreadsheets_path = spreadsheets_path
        self.spreadsheets = {}
        self.locks = {} # A lock for each spreadsheet

        
    def __logging(self):
        """Set up logging."""
        
        logging.basicConfig(format='%(asctime)s:%(levelname)s:%(message)s',
                            datefmt='%Y%m%d %H:%M:%S',
                            filename=self.log_file,
                            level=logging.INFO)

        print('Logging to: ' + self.log_file)



    def __start_soffice(self):

        logging.info('Starting the soffice process.')
        command = '/usr/bin/soffice --accept="pipe,name=' + self.soffice_pipe +';urp;"\
        --norestore --nologo --nodefault --headless'

        logfile = open(self.soffice_log, "w")
        self.soffice_process = subprocess.Popen(command, shell=True, stdout=logfile, stderr=logfile)

        
    def __connect_to_soffice(self):
        """Make a connection to soffice and fail if it can not connect."""
        
        MAX_ATTEMPTS = 60

        attempt = 0
        while 1:
            if attempt > MAX_ATTEMPTS: # soffice process isin't coming up
                raise RuntimeError("Could not connect to soffice process.")

            try:
                self.soffice = pyoo.Desktop(pipe=SOFFICE_PIPE)
                logging.info("Connected to soffice.")
                break

            except OSError:
                attempt += 1
                sleep(1)

    def __start_threaded_tcp_server(self):
        """Set up and start the TCP threaded server to handle incomming 
        requests.
        """

        logging.info('Starting spreadsheet_server.')
        
        try:
            self.server = ThreadedTCPServer(
                (self.soffice_host, self.soffice_port),
                ThreadedTCPRequestHandler
            )
            
        except OSError:
            print("Error: The port is in use. Maybe the server is already running?")
            exit()


        self.server.spreadsheets = self.spreadsheets
        self.server.locks = self.locks

        # Start the main server thread. This server thread will start a
        # new thread to handle each client connection.

        self.server_thread = threading.Thread(
            target=self.server.serve_forever)
        
        self.server_thread.daemon = False # Gracefully stop child threads
        self.server_thread.start()

        logging.info("Server thread running. Waiting on connections...")


    def __start_monitor_thread(self):
        """This thread monitors the SPREADSHEETS directory to add or remove.
        """
        
        self.monitor_thread = MonitorThread(self.spreadsheets, self.locks,
                                       self.soffice, self.spreadsheets_path,
                                       self.monitor_frequency)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()


    def stop(self):
        """Stop all the threads and shutdown LibreOffice."""

        # Stop the monitor thread
        self.monitor_thread.stop_thread()
        self.monitor_thread.join()

        # Stop the ThreadedTCPServer
        self.server.shutdown()

        # Terminate the soffice.bin process
        self.soffice_process.send_signal(SIGTERM)

    
    def run(self):
        self.__logging()
        self.__start_soffice()
        self.__connect_to_soffice()
        self.__start_threaded_tcp_server()
        self.__start_monitor_thread()

            
if __name__ == "__main__":

    print('Starting spreadsheet_server...')

    spreadsheet_server = SpreadsheetServer()
    try:
        spreadsheet_server.run()
        print('Up and listening for connections!')
        while True: sleep(100)
        
    except (KeyboardInterrupt, SystemExit):
        print("Shutting down server...")
        spreadsheet_server.stop()
        

