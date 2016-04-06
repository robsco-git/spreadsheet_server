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
    def handle(self):
        def send(msg):
            # Send utf-8 encoded bytes of the json encoded strings
            # msg encoded with json -> bytes encoded in utf-8
            self.request.sendall(bytes(json.dumps(msg), "utf-8"))
            logging.debug("Sent: " + json.dumps(msg))
        def receive():
            # convert the received utf-8 bytes into a string -> load the object via json
            recv = self.request.recv(4096)
            if recv == b'':
                # connection is closed
                return False
            received = json.loads(str(recv, encoding="utf-8"))
            logging.debug("Received: " + str(received))
            return received
                
        try:
            # Handle first request to server and check that it adheres to the protocol
            data = receive()
            if (data[0] == "SPREADSHEET"): # correct protocol
                this_con = SpreadsheetConnection(server.spreadsheets[data[1]], server.locks[data[1]])
                send("OK")
                # lock the spreadsheet
                this_con.lock_spreadsheet()
                # raise Exception("Manual Error")
                # Run main loop for all communication
                while True:
                    data = receive()
                    if data == False:
                        # connection lost
                        break
                    elif data[0] == "SET":
                        # try:
                        this_con.set_cells(data[1], data[2], data[3])
                        send("OK")
                        # except ValueError:
                        #     send("ERROR")
                    elif data[0] == "GET":
                        cells = this_con.get_cells(data[1], data[2])

                        if cells != False:
                            send(cells)
                        else:
                            send("ERROR")
                    elif data[0] == "SAVE":
                        this_con.save_spreadsheet(data[1])
                        send("OK")
        except:
            logging.debug(traceback.print_exc())
            logging.debug("Connection to client lost")
        finally:
            # make sure to unlock the spreadsheet
            try:
                if this_con.lock.locked:
                    this_con.unlock_spreadsheet()
            except UnboundLocalError:
                # this_con was never created
                pass
                
            # close the socket
            try:
                self.request.shutdown(SHUT_RDWR)
            except OSError:
                # client already disconnected
                pass
                
            self.request.close()
        
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
