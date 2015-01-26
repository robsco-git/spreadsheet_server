import socket
import json
import traceback

class SpreadsheetClient:
    def __init__(self, ip, port, workbook):
        try:
            self.sock = self.connect(ip, port)
        except socket.error:
            raise Exception("Could not connect to server!")
        else:
            self.set_workbook(workbook)
        
    def connect(self, ip, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        return sock
                
    def set_workbook(self, workbook):
        self._send(["WORKBOOK", workbook])
        if self._receive() == "OK":
            return True
        else:
            raise Exception("Workbook could not be set!")

    def set_cells(self, sheet, cell_range, data):
        self._send(["SET", sheet, cell_range, data])
        if self._receive() == "OK":
            return True
        else:
            raise Exception("Could not set cells!")

    def get_cells(self, sheet, cell_range):
        self._send(["GET", sheet, cell_range])
        return self._receive()
        
    def save_workbook(self, filename):
        self._send(["SAVE", filename])
        return self._receive()
        
    def _send(self, msg):
        try:
            # endoce msg into json then send over the socket
            self.sock.sendall(json.dumps(msg, encoding='utf-8'))
            # print(msg)
        except:
            raise Exception("Could not send message to server")
            traceback.print_exc()
            print("Connection error")

    def _receive(self):
        # convert the received utf-8 bytes into a string -> load the object via json
        recv = self.sock.recv(4096)
        if recv == b'':
            # connection is closed
            raise Exception("Connection to server closed!")
        
        received = json.loads(recv, encoding="utf-8")
        # print("Received: " + str(received))
        return received
            
    def disconnect(self):
        try:
            self.sock.shutdown(socket.SHUT_RDWR)
        except socket.error:
            # client already disconnected
            pass
        self.sock.close()
        
