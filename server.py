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
from signal import SIGTERM
import os
import fileinput
import sys


SOFFICE_BINARY = "soffice.bin"
LOG_FILE = './log/server.log'
SOFFICE_LOG = './log/soffice.log'
HOST, PORT = "localhost", 5555
SOFFICE_PIPE = "soffice_headless"
SPREADSHEETS_PATH = "./spreadsheets"
MONITOR_FREQ = 5 # In seconds
SAVE_PATH = "./saved_spreadsheets/"

class SpreadsheetServer:

    def __init__(self, log_file=LOG_FILE, soffice_log=SOFFICE_LOG,
                 host=HOST, port=PORT,
                 soffice_pipe=SOFFICE_PIPE,
                 spreadsheets_path=SPREADSHEETS_PATH,
                 monitor_frequency=MONITOR_FREQ,
                 reload_on_disk_change=True,
                 ask_kill=False,
                 save_path=SAVE_PATH):
        
        self.log_file = log_file # Where 'logging' logs to

        # Where the output from LibreOffice is logged to
        self.soffice_log = soffice_log

        self.host = host # The address on which the server is listening
        self.port = port # The port on which the server is listening

        # The name of the pipe set up by LibreOffice that pyoo will connect to.
        self.soffice_pipe = soffice_pipe

        # The frequency, in seconds, at which the directory containing the
        # spreadsheets is polled.
        self.monitor_frequency = monitor_frequency

        # Weather or not to close and open a spreadsheet with the file changes on
        # disk
        self.reload_on_disk_change = reload_on_disk_change

        # Whether or not to interactively ask the user if they want to kill an
        # existing LibreOffice process. 
        self.ask_kill = ask_kill

        # Where to save the spreadsheets when requested by the client.
        self.save_path = save_path

        # Where to look for spreadsheets to load.
        self.spreadsheets_path = spreadsheets_path

        self.spreadsheets = {} # Each pyoo spreadsheet object.
        self.locks = {} # A lock for each spreadsheet.
        self.hashes = {} # A hash of the file contents for each spreadsheet.

        
    def __logging(self):
        """Set up logging."""
        
        logging.basicConfig(format='%(asctime)s:%(levelname)s:%(message)s',
                            datefmt='%Y%m%d %H:%M:%S',
                            filename=self.log_file,
                            level=logging.INFO)


    def __start_soffice(self):

        def get_pid(name):
            return map(int,subprocess.check_output(["pidof",name]).split())

        
        def killall(pids):
            logging.warning('Killing existing LibreOffice process')
            for pid in pids:
                os.kill(pid, SIGTERM)
        

        # Check for a already running LibreOffice process
        try:
            pids = get_pid(SOFFICE_BINARY)

            if self.ask_kill:
                ask_str = "LibreOffice is already running. Would you like to kill it? (Y/n): "

                while True:

                    answer = input(ask_str)
                    answer = answer.lower()
                    
                    if answer not in ['y', 'n', '']:
                        Print("Please respond with 'y' or 'n'.")
                    else:
                        break

                if answer in ['y', '']:
                    killall(pids)
                else:
                    print("Goodbye!")
                    exit()
                    
            else:
                killall(pids)
                
        except subprocess.CalledProcessError:
            # There is no soffice.bin process
            pass

        # Use which to get the binary location
        
        logging.info('Starting the soffice process.')
        command = '/usr/bin/soffice --accept="pipe,name=' + self.soffice_pipe +';urp;"\
        --norestore --nologo --nodefault --headless'

        self.logfile = open(self.soffice_log, "w")
        self.soffice_process = subprocess.Popen(
            command, shell=True, stdout=self.logfile, stderr=self.logfile)

        
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

            except IOError:
                print("IOError in connection to soffice")
                attempt += 1
                sleep(1)

            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(exc_type, fname, exc_tb.tb_lineno)
                
                
    def __start_threaded_tcp_server(self):
        """Set up and start the TCP threaded server to handle incomming 
        requests.
        """

        logging.info('Starting spreadsheet_server.')
        
        try:
            self.server = ThreadedTCPServer(self.save_path,
                (self.host, self.port),
                ThreadedTCPRequestHandler
            )
            
        except OSError:
            print("Error: The port is in use. Maybe the server is already running?")
            exit()


        self.server.spreadsheets = self.spreadsheets
        self.server.locks = self.locks
        self.server.hashes = self.hashes

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
        
        self.monitor_thread = MonitorThread(
            self.spreadsheets, self.locks, self.hashes, self.soffice,
            self.spreadsheets_path, self.monitor_frequency,
            self.reload_on_disk_change)
        
        self.monitor_thread.daemon = True
        self.monitor_thread.start()


    def __stop_monitor_thread(self):
        """Stop the monitor thread."""
        
        self.monitor_thread.stop_thread()
        self.monitor_thread.join()


    def __stop_threaded_tcp_server(self):
        """Stop the ThreadedTCPServer."""

        try:
            self.server.server_close()
            self.server.shutdown()
        except AttributeError:
            # The server was never set up
            pass
        

    def __kill_libreoffice(self):
        """Terminate the soffice.bin process."""

        self.soffice_process.send_signal(SIGTERM)        
        

    def __close_logfile(self):
        """Close the logfile."""
        
        self.logfile.close()
        

    def stop(self):
        """Stop all the threads and shutdown LibreOffice."""

        self.__stop_monitor_thread()
        self.__stop_threaded_tcp_server()
        self.__kill_libreoffice()
        self.__close_logfile()

    
    def run(self):
        self.__logging()
        self.__start_soffice()
        self.__connect_to_soffice()
        self.__start_monitor_thread()
        self.__start_threaded_tcp_server()

            
if __name__ == "__main__":

    print('Starting spreadsheet_server...')

    spreadsheet_server = SpreadsheetServer(ask_kill=True)
    try:
        print('Logging to: ' + spreadsheet_server.log_file)
        spreadsheet_server.run()
        print('Up and listening for connections!')
        while True: sleep(100)
        
    except KeyboardInterrupt:
        print("Shutting down server. Please wait...")
        spreadsheet_server.stop()
        

