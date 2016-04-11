# This file is part of calc_server.

# calc_server is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# calc_server is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with calc_server.  If not, see <http://www.gnu.org/licenses/>.


import random
from client_python2 import SpreadsheetClient

if __name__ == "__main__":
    example_spreadsheet = "example.ods"
    try:
        client = SpreadsheetClient("localhost", 5555, example_spreadsheet)
    except:
        print("Could not connect to server")
    else:
        try:
            cell_range = "A1:C3"
            print "Cells " + cell_range
            all_values = client.get_cells('Sheet1', cell_range)
            print all_values

        except:
            print("Error in connection")
        finally:
            client.disconnect()
