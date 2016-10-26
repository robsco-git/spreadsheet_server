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

import socketserver
import json
from socket import SHUT_RDWR
import logging
from connection import SpreadsheetConnection


class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    pass


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
            self.server.spreadsheets[data[1]], self.server.locks[data[1]])
        
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

            elif data[0] == "GET_SHEETS":
                sheet_names = self.con.get_sheet_names()
                if sheet_names != False:
                    self.__send(sheet_names)
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

        
