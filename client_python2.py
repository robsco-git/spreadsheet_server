import socket
import json
import traceback

class SpreadsheetClient:
    def __init__(self, ip, port, spreadsheet):
        try:
            self.sock = self.connect(ip, port)
        except socket.error:
            raise Exception("Could not connect to server!")
        else:
            self.set_spreadsheet(spreadsheet)
        
    def connect(self, ip, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        return sock
                
    def set_spreadsheet(self, spreadsheet):
        self.__send(["SPREADSHEET", spreadsheet])
        if self.__receive() == "OK":
            return True
        else:
            raise Exception("Spreadsheet could not be set!")

    def set_cells(self, sheet, cell_range, data):
        self.__send(["SET", sheet, cell_range, data])
        if self.__receive() == "OK":
            return True
        else:
            raise Exception("Could not set cells!")

    def get_cells(self, sheet, cell_range):
        self.__send(["GET", sheet, cell_range])
        return self.__receive()
        
    def save_spreadsheet(self, filename):
        self.__send(["SAVE", filename])
        return self.__receive()
        
    def __send(self, msg):
        try:
            # endoce msg into json then send over the socket
            self.sock.sendall(json.dumps(msg, encoding='utf-8'))
            # print(msg)
        except:
            raise Exception("Could not send message to server")
            traceback.print_exc()
            print("Connection error")

    def __receive(self):
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
        
