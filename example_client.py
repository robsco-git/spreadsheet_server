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

from client import SpreadsheetClient

if __name__ == "__main__":
    """This script shows how the differnet functions exposed by client.py can be
    used."""
    
    EXAMPLE_SPREADSHEET = "example.ods"
    SHEET_NAME = "Sheet1"
    sc = SpreadsheetClient("localhost", 5555, EXAMPLE_SPREADSHEET)
    
    # Set a cell value
    sc.set_cells(SHEET_NAME, "A1", 5)

    # Retrieve a cell value.
    cell_value = sc.get_cells(SHEET_NAME, "C3")
    print(cell_value)

    # Set a one dimensional cell range.
    # Cells are set using the format: [[A1, B1], [A2, B2], [A3, B3]]
    cell_values = [[6, 5], [4, 3], [2, 1]]
    sc.set_cells(SHEET_NAME, "A1:B3", cell_values)

    # Retrieve one dimensional cell range.
    cell_values = sc.get_cells(SHEET_NAME, "C1:C3")
    print(cell_values)

    # Set a two dimensional cell range.
    # Cells are set using the format: [[A1, B1, C1], [A2, B2, C2], [A3, B3, C3]]
    cell_values = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
    sc.set_cells(SHEET_NAME, "A1:C3", cell_values)

    # Retrieve a two dimensional cell range.
    cell_values = sc.get_cells(SHEET_NAME, "A1:C3")
    print(cell_values)

    # Save a spreadsheet - it will save into ./saved_spreadsheets
    sc.save_spreadsheet(EXAMPLE_SPREADSHEET)
    
    sc.disconnect()
