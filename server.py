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
        # Unlock the spreadsheet
        try:
            if self.con.lock.locked:
                self.con.unlock_spreadsheet()
        except UnboundLocalError:
            # con was never created
            pass
            
        # Close the connection to the client
        try:
            self.request.shutdown(SHUT_RDWR)
        except OSError:
            # client already disconnected
            pass

        self.request.close()


    def __main_loop(self):
        while True:
            data = self.__receive()
            
            if data == False:
                # The connection was lost
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
        """Not sure of the description here yet...
        
        """

        self.__make_connection()
        self.__main_loop()
        self.__close_connection()

        
class MonitorThread(threading.Thread):
    """ Monitors the spreadsheet directory for changes """
    def run(self):
        while True:
            docs = [ f for f in listdir(SPREADSHEETS_PATH) if isfile(join(SPREADSHEETS_PATH,f)) ]
            # check for removed spreadsheets
            removed_spreadsheets = []
            for key, value in server.spreadsheets.items():
                removed = True
                for doc in docs:
                    if key == doc:
                        removed = False
                        break
                if removed:
                    removed_spreadsheets.append(key)

            for doc in removed_spreadsheets:
                logging.info("Removing " + doc)
                server.locks[doc].acquire()
                server.spreadsheets[doc].close()
                server.spreadsheets.pop(doc, None)
                server.locks.pop(doc, None)

            # check for new spreadsheets
            for doc in docs:
                if doc[0] != '.':
                    found = False
                   
                    for key, value in server.spreadsheets.items():
                        if doc == key:
                            found = True
                            break
                    if found == False:
                        logging.info("Loading " + doc)
                        server.spreadsheets[doc] = soffice.open_spreadsheet(SPREADSHEETS_PATH + "/" + doc)
                        server.locks[doc] = threading.Lock()
            sleep(MONITOR_THREAD_FREQ)
        
class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    pass

if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(format='%(asctime)s:%(levelname)s:%(message)s',
                        datefmt='%Y%m%d %H:%M:%S',
                        filename='./log/server.log',
                        level=logging.DEBUG)
    logging.info('Started')
    logging.info('Killing any current soffice process')
    # Kill any currently running soffice process
    subprocess.call(['killall', 'soffice.bin'])

    subprocess.call(['rm', '/tmp/OSL_PIPE_1001_soffice_headless'])

    logging.info('Starting soffice process')
    # Start the headless soffice process on the server
    command = '/usr/bin/soffice --accept="pipe,name=' + SOFFICE_PIPE + ';urp;" --norestore --nologo --nodefault --headless'
    logfile = open("./log/soffice.log", "w")
    subprocess.Popen(command, shell=True, stdout=logfile, stderr=logfile)

    # Make a connection to soffice
    attempts = 0
    while 1:
        try:
            if attempts > 300: # soffice process isin't coming up
                exit()
            soffice = pyoo.Desktop(pipe=SOFFICE_PIPE)
            logging.info("Connected to soffice.")
            break
        except NoConnectException:
            logging.debug('Could not connect to soffice process. Attempt no: ' + str(attempts+1))
            attempts += 1
            sleep(1)

    # Start server initialisation
    HOST, PORT = "localhost", 5555

    server = ThreadedTCPServer((HOST, PORT), ThreadedTCPRequestHandler)
    
    # create documents for each file in ./spreadsheets
    docs = [ f for f in listdir(SPREADSHEETS_PATH) if isfile(join(SPREADSHEETS_PATH,f)) ]
    server.spreadsheets = {}
    server.locks = {} # a lock for each spreadsheet
    for doc in docs:
        if doc[0] != '.' :
            logging.info("Loading " + doc)
            server.spreadsheets[doc] = soffice.open_spreadsheet(SPREADSHEETS_PATH + "/" + doc)
            server.locks[doc] = threading.Lock()

    # Start a thread with the server -- that thread will then start one
    # more thread for each request
    server_thread = threading.Thread(target=server.serve_forever)
    
    # Don't exit the server thread when the main thread terminates
    server_thread.daemon = False
    logging.info('Server starting')
    server_thread.start()

    # Update spreadsheets thread
    monitor_thread = MonitorThread()
    monitor_thread.daemon = True
    monitor_thread.start()
    
    logging.info("Server thread running. Waiting on connections...")
