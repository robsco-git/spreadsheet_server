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

import socket
import json
import traceback
import struct
import select

IP, PORT = "localhost", 5555

TIMEOUT = 10


class SpreadsheetClient:
    def __init__(self, spreadsheet, ip=IP, port=PORT):
        try:
            self.sock = self.__connect(ip, port)
        except socket.error:
            raise RuntimeError("Could not connect to the server.")
        else:
            self.__set_spreadsheet(spreadsheet)

    def __connect(self, ip, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((ip, port))
        return sock

    def __set_spreadsheet(self, spreadsheet):
        self.__send(["SPREADSHEET", spreadsheet])
        received = self.__receive()

        if received == "NOT FOUND" or received != "OK":
            self.disconnect()
            raise RuntimeError("The requested spreadsheet was not found.")

    def set_cells(self, sheet, cell_ref, data):
        """Set the value(s) for a single cell or a cell range.

        'sheet' is either a 0-based index or the string name of the sheet.
        'cell_ref' is a LibreOffice style cell reference. eg. "A1" or "D7:G42".

        For a single cell, 'data' is a single string, int or float value.

        For a one dimensional (only horizontal or only vertical) range of
        cells, 'data' is a list. For a two dimensional range of cells, 'data'
        is a list of lists. For example setting the 'cell_ref' "A1:C3"
        requires 'data' of the format:
        [[A1, B1, C1], [A2, B2, C2], [A3, B3, C3]].
        """

        self.__send(["SET", sheet, cell_ref, data])

        received = self.__receive()
        if type(received) == dict:
            # The server is retuning an error
            raise RuntimeError(received["ERROR"])

    def get_sheet_names(self):
        """Returns a list of all sheet names in the workbook."""

        self.__send(["GET_SHEETS"])
        sheet_names = self.__receive()

        if sheet_names == "ERROR":
            raise Exception("Could not retrieve sheet names.")

        return sheet_names

    def get_cells(self, sheet, cell_ref):
        """Get the value of a single cell or a cell range from the server
        and return it or them.

        'sheet' is either a 0-based index or the string name of the sheet.
        'cell_ref' is what one would use in LibreOffice Calc. Eg. "ABC945" or
        "A3:F75".

        A single cell is returned for a single value.
        A list of cell values is returned for a one dimensional range of cells.
        A list of lists is returned for a two dimensional range of cells.
        """

        self.__send(["GET", sheet, cell_ref])
        cells = self.__receive()

        if type(cells) == dict:
            # The server is retuning an error
            raise RuntimeError(cells["ERROR"])

        return cells

    def save_spreadsheet(self, filename):
        """Save the spreadsheet in its current state on the server. The
        server determines where it is saved."""

        self.__send(["SAVE", filename])
        return self.__receive()

    def __send(self, msg):
        """Encode msg into json and then send it over the socket."""

        json_msg = json.dumps(msg)
        json_msg = bytes(json_msg, "utf-8")

        # Prepend the length of the string to the meg
        json_msg = struct.pack(">I", len(json_msg)) + json_msg

        try:
            self.sock.sendall(json_msg)
        except:  # noqa
            traceback.print_exc()
            raise Exception("Could not send message to server")

    def __receive(self):
        """Receive a message from the client, convert the received utf-8
        bytes into a string then decode if from json."""

        raw_msg_length = self.__receive_length(4)
        if not raw_msg_length:
            return False
        msg_length = struct.unpack(">I", raw_msg_length)[0]

        recv = self.__receive_length(msg_length)

        if recv == b"":
            # The connection has been closed.
            raise Exception("Connection to server closed!")

        received = str(recv, encoding="utf-8")
        received = json.loads(received)

        return received

    def __receive_length(self, length):
        """Receive length number of bytes from the client."""

        data = b""
        while len(data) < length:

            ready = select.select([self.sock], [], [], TIMEOUT)

            if ready[0]:

                packet = self.sock.recv(length - len(data))
                if not packet:
                    return b""
                data += packet

            else:
                # Did not recieve on the socket within the timeout
                return b""

        return data

    def disconnect(self):
        """Disconnect from the server."""

        try:
            self.sock.shutdown(socket.SHUT_RDWR)
        except socket.error:
            # The client has already disconnected
            pass

        self.sock.close()
