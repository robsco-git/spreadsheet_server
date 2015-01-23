import socket
import json
import traceback
import random

class Client:
    def __init__(self, ip, port, workbook):
        self.sock = self.connect(ip, port)
        self.set_workbook(workbook)

    def connect(self, ip, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        return sock
                
    def set_workbook(self, workbook):
        self._send(["WORKBOOK", workbook])
        return self._receive()

    # def lock_workbook(self):
    #     self._send(["LOCK"])
    #     return self._receive()

    def set_cells(self, sheet, cell_range, data):
        self._send(["SET", sheet, cell_range, data])
        if self._receive() == "OK":
            return True
        else:
            return False

    def get_cells(self, sheet, cell_range):
        self._send(["GET", sheet, cell_range])
        return self._receive()
        
    # def unlock_workbook(self):
    #     self._send(["UNLOCK"])
    #     return self._receive()

    def save_workbook(self):
        self._send(["SAVE"])
        return seld._receive()
        
    def _send(self, msg):
        try:
            # endoce msg into json then send over the socket
            self.sock.sendall(json.dumps(msg, encoding='utf-8'))
            print(msg)
        except:
            traceback.print_exc()
            print("Connection error")

    def _receive(self):
        # convert the received utf-8 bytes into a string -> load the object via json
        recv = self.sock.recv(4096)
        if recv == b'':
            # connection is closed
            return False
        received = json.loads(recv, encoding="utf-8")
        print("Received: " + str(received))
        return received
            
    def disconnect(self):
        self.sock.shutdown(socket.SHUT_RDWR)
        self.sock.close()

if __name__ == "__main__":
    testing_doc = "20141209 HQA testsheet 1 (GS) BM.xlsx"
    #testing_doc = "range_test.xlsx"
    client = Client("localhost", 5555, testing_doc)
    for x in range(0, 1):
        
        options = client.get_cells('Input', "B4:B94")
        rand_opt = random.choice(options)
        print(rand_opt)
        client.set_cells("Input", "B2", rand_opt)
        print(client.get_cells("Calcsheet", "K7:K21"))
    client.disconnect()
