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

import random
from client_python2 import SpreadsheetClient

if __name__ == "__main__":
    example_spreadsheet = "example.ods"
    client = SpreadsheetClient("localhost", 5555, example_spreadsheet)
    
    cell_range = "A1:C3"
    print "Cells " + cell_range
    all_values = client.get_cells('Sheet1', cell_range)
    print all_values
    client.disconnect()

        
