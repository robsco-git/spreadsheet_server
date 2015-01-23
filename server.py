# a change for git

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
import con
import json
import traceback

SOFFICE_PIPE = "soffice_headless"
WORKBOOKS_PATH = "/home/ubuntu/projects/pyoo_server/workbooks"
MONITOR_THREAD_FREQ = 60 # In seconds

class ThreadedTCPRequestHandler(socketserver.BaseRequestHandler):
    def handle(self):
        print("New connection")
        def send(msg):
            # Send utf-8 encoded bytes of the json encoded strings
            # msg encoded with json -> bytes encoded in utf-8
            self.request.sendall(bytes(json.dumps(msg), "utf-8"))
            print("Sent: " + json.dumps(msg))
        def receive():
            # convert the received utf-8 bytes into a string -> load the object via json
            recv = self.request.recv(4096)
            if recv == b'':
                # connection is closed
                return False
            received = json.loads(str(recv, encoding="utf-8"))
            print("Received: " + str(received))
            return received
                
        try:
            # Handle first request to server and check that it adheres to the protocol
            data = receive()
            if (data[0] == "WORKBOOK"): # correct protocol
                this_con = con.ConnectionHandler(server.workbooks[data[1]], server.locks[data[1]])
                send("OK")
                # lock the workbook
                this_con.lock_workbook()
                raise Exception("Manual Error")
                # Run main loop for all communication
                while True:
                    data = receive()
                    if data == False:
                        # connection lost
                        break
                    # elif data[0] == "LOCK":
                    #     this_con.lock_workbook()
                    #     send("OK")
                    elif data[0] == "SET":
                        if this_con.set_cells(data[1], data[2], data[3]):
                            send("OK")
                        else:
                            send("ERROR")
                    elif data[0] == "GET":
                        cells = this_con.get_cells(data[1], data[2])
                        if cells != False:
                            send(cells)
                        else:
                            send("ERROR")
                    # elif data[0] == "UNLOCK":
                    #     if this_con.unlock_workbook():
                    #         send("OK")
                    #     else:
                    #         send("ERROR")
                    elif data[0] == "SAVE":
                        this_con.save_workbook()
                        send("OK")
        except:
            traceback.print_exc()
            print("Connection to client lost")
        finally:
            # make sure to unlock the workbook
            try:
                if this_con.lock.locked:
                    this_con.unlock_workbook()
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
    """ Monitors the workbook directory for changes """
    def run(self):
        while True:
            docs = [ f for f in listdir(WORKBOOKS_PATH) if isfile(join(WORKBOOKS_PATH,f)) ]
            # check for removed workbooks
            removed_workbooks = []
            for key, value in server.workbooks.items():
                removed = True
                for doc in docs:
                    if key == doc:
                        removed = False
                        break
                if removed:
                    removed_workbooks.append(key)

            for doc in removed_workbooks:
                print("Removing " + doc)
                server.locks[doc].acquire()
                server.workbooks[doc].close()
                server.workbooks.pop(doc, None)
                server.locks.pop(doc, None)

            # check for new workbooks
            for doc in docs:
                if doc[0] != '.':
                    found = False
                   
                    for key, value in server.workbooks.items():
                        if doc == key:
                            found = True
                            break
                    if found == False:
                        print("Loading " + doc)
                        server.workbooks[doc] = soffice.open_spreadsheet(WORKBOOKS_PATH + "/" + doc)
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
    #command = shlex.split(command)
    logfile = open("./log/soffice.log", "w")
    subprocess.Popen(command, shell=True, stdout=logfile, stderr=logfile)

    # Make a connection to soffice
    attempts = 0
    while 1:
        try:
            if attempts > 300: # soffice process isin't coming up
                exit()
            soffice = pyoo.Desktop(pipe=SOFFICE_PIPE)
            print("Connected to soffice.")
            break
        except NoConnectException:
            logging.debug('Could not connect to soffice process. Attempt no: ' + str(attempts+1))
            attempts += 1
            sleep(1)

    # Start server initialisation
    HOST, PORT = "localhost", 5555

    server = ThreadedTCPServer((HOST, PORT), ThreadedTCPRequestHandler)
    
    # create documents for each file in ./workbooks
    docs = [ f for f in listdir(WORKBOOKS_PATH) if isfile(join(WORKBOOKS_PATH,f)) ]
    server.workbooks = {}
    server.locks = {} # a lock for each workbook
    for doc in docs:
        if doc[0] != '.' :
            print("Loading " + doc)
            server.workbooks[doc] = soffice.open_spreadsheet(WORKBOOKS_PATH + "/" + doc)
            server.locks[doc] = threading.Lock()


    # Start a thread with the server -- that thread will then start one
    # more thread for each request
    server_thread = threading.Thread(target=server.serve_forever)
    
    # Don't exit the server thread when the main thread terminates
    server_thread.daemon = False
    logging.info('Server starting')
    server_thread.start()

    # Update workbooks thread
    monitor_thread = MonitorThread()
    monitor_thread.daemon = True
    monitor_thread.start()
    
    print("Server loop running in thread:", server_thread.name)

    #server.shutdown()
