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
from time import sleep


class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):

    def __init__(self, save_path, *args, **kwargs):
        self.save_path = save_path
        socketserver.TCPServer.__init__(self, *args, **kwargs)


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

        def protocol_error():
            # Invalid connection protocol
            self.logging.error(
                "Client attempted to connect using and invalid protocol."
            )
            self.__send("PROTOCOL ERROR")
            self.__close_connection()
            return False
        
        data = self.__receive()

        if type(data) != list:
            return protocol_error()

        if len(data) != 2:
            return protocol_error()
        
        if (data[0] != "SPREADSHEET"):
            return protocol_error()
        
        # If there is a KeyError when looking up the spreadsheets name, wait
        # a bit and try again

        max_attempts = self.server.monitor_frequency + 1
        attempt = 0
        
        while 1:
            if attempt >= max_attempts: # soffice process isin't coming up
                logging.debug("Waited too long for spreadsheet")
                # We can assume the spreadsheet does not exist
                self.__send("NOT FOUND")
                self.__close_connection()
                return False

            try:
                self.con = SpreadsheetConnection(
                    self.server.spreadsheets[data[1]],
                    self.server.locks[data[1]],
                    self.server.save_path
                )
                break
                
            except KeyError:
                pass

            attempt += 1
            logging.debug("Waiting for spreadsheet")
            sleep(1)

        # If the spreadsheet was sucessfully connected to
        if attempt != max_attempts:
            self.__send("OK")
            self.con.lock_spreadsheet()
            return True


    def __close_connection(self):
        """Unlock the spreadsheet and close the connection to the client."""
        
        try:
            if self.con.lock.locked:
                self.con.unlock_spreadsheet()
        except (UnboundLocalError, AttributeError):
            # con was never created.
            pass
            
        try:
            self.request.shutdown(SHUT_RDWR)
        except OSError:
            # The client has already disconnected.
            pass

        logging.debug("Closing socket for ThreadedTCPRequestHandler")
        self.request.close()


    def __main_loop(self):
        while True:
            data = self.__receive()

            if data == False:
                # The connection has been lost.
                break
            
            elif data[0] == "SET":
                try:
                    self.con.set_cells(data[1], data[2], data[3])
                except ValueError as e:
                    self.__send({"ERROR": str(e)})
                else:
                    self.__send("OK")
                
            elif data[0] == "GET":
                try:
                    cells = self.con.get_cells(data[1], data[2])
                except ValueError as e:
                    self.__send({"ERROR": str(e)})
                else:
                    self.__send(cells)

            elif data[0] == "GET_SHEETS":
                sheet_names = self.con.get_sheet_names()
                self.__send(sheet_names)
                    
            elif data[0] == "SAVE":
                self.con.save_spreadsheet(data[1])
                self.__send("OK")

    
    def handle(self):
        """Make a connection to the client, run the main protocol loop and
        close the connection.
        """

        if self.__make_connection():
            self.__main_loop()
            self.__close_connection()

