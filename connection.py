import logging
from math import pow
from com.sun.star.uno import RuntimeException
import traceback

class SpreadsheetConnection:
    """Handles connections to the spreadsheets opened by soffice (LibreOffice).
    """
    
    def __init__(self, spreadsheet, lock):
        self.spreadsheet = spreadsheet
        self.lock = lock


    def lock_spreadsheet(self):
        """Lock the spreadsheet.
        
        The getting and setting cell functions rely on a given spreadsheet 
        being locked. This insures simulations requests to the same spreadsheet
        do not interfere with one another.
        """
        
        self.lock.acquire()


    def unlock_spreadsheet(self):
        """ Unlock the spreadsheet and return a 'success' boolean."""
        
        try:
            self.lock.release()
            return True
        except RuntimeError:
            return False


    def _get_xy_index(self, cell_ref):
        chars = [c for c in cell_ref if c.isalpha()]
        nums = [n for n in cell_ref if n.isnumeric()]

        alpha_index = 0
        for i, c in enumerate(chars):
            # The base index for this character is in the range of 0 to 25
            c_index = ord(c.upper()) - 65

            if i == len(chars) - 1:
                # The simple case of the least significant character.
                # Eg. The 'J' in 'AMJ'.
                alpha_index += c_index

            else:
                # The index for additional characters to the left are
                # calculated c_index * 26^(n) where n is the characters
                # position to the left.
                alpha_index += c_index * pow(26, len(chars)-i-1)

        num_index = int(''.join(nums)) - 1 # zero-based
        return alpha_index, num_index


    def _is_single_cell(self, cell_ref):
        if len(cell_ref.split(':')) == 1:
            return True
        return False


    def _check_single_cell(self, cell_ref):
        if not self._is_single_cell(cell_ref):
            raise ValueError(
                "Expected a single cell reference. A cell range was given.")


    def _cell_to_index(self, cell_ref):
        """Convert a spreadsheet style single cell or cell reference, to a zero 
        based numerical index.

        'cell_ref' is what one would use in LibreOffice Calc. Eg. "ABC945".

        Returned is: {"row_index": int, "column_index": int}.
        """

        
        alpha_index, num_index = self._get_xy_index(cell_ref)
        return {"row_index": num_index, "column_index": alpha_index}

        
    def _cell_range_to_index(self, cell_ref):
        """Convert a spreadsheet style range reference to zero-based numerical
        indecies that describe the start and end points of the cell range.

        'cell_ref' is what one would use in LibreOffice Calc. Eg. "A1" or
        "A1:D6".

        Returned is: {"row_start": int, "row_end": int, 
        "column_start": int, "column_end": int}
        """

        left_ref, right_ref = cell_ref.split(':')

        left_alpha_index, left_num_index = self._get_xy_index(left_ref)
        right_alpha_index, right_num_index = self._get_xy_index(right_ref)

        return {"row_start": left_num_index,
                "row_end": right_num_index,
                "column_start": left_alpha_index,
                "column_end": right_alpha_index}


    def _check_for_lock(self):
        if not self.lock.locked():
            raise RuntimeError(
                "Lock for this spreadsheet has not been aquired.")


    def _check_numeric(self, value):
        """Convert a string representation of a number to a float."""
        try:
            return float(value)
        except ValueError:
            return value


    def _check_list(self, data):
        if not isinstance(data, list):
            raise ValueError("Expecting list type.")


    def _check_1D_list(self, data):
        check_list(data)
        if isinstance(data[0], list):
            raise ValueError("Got 2D list when expecting 1D list.")

        for x, cell in enumerate(data):
            data[x] = check_numeric(cell)


    def set_cells(self, sheet, cell_ref, value):
        """Sets the value(s) of a single cell or a cell range. This can be used
        when it is not known if 'cell_ref' refers to a single cell or a range

        See 'set_cell' and 'set_cell_range' for more information.
        """
        
        if self._is_single_cell(cell_ref):
            self.set_cell(sheet, cell_ref, value)
        else:
            self.set_cell_range(sheet, cell_ref, value)
            

    def set_cell(self, sheet, cell_ref, value):
        """Set the value of a single cell.

        'sheet' is either a 0-based index or the string name of the sheet. 
        'cell_ref' is a LibreOffice style cell reference. eg. "A1".
        'value' is a single string, int or float value.
        """

        self._check_single_cell(cell_ref)
        
        self._check_for_lock()
        
        r = self._cell_to_index(cell_ref)
        sheet = self.spreadsheet.sheets[sheet]

        if isinstance(value, list):
            raise ValueError("Expectin a single cell. \
            A list of cells was given.")

        value = self._check_numeric(value)
        sheet[r["row_index"], r["column_index"]].value = value


    def set_cell_range(self, sheet, cell_ref, data):
        """Set the values for a cell range.

        'sheet' is either a 0-based index or the string name of the sheet. 
        'cell_ref' is a LibreOffice style cell reference. eg. "D7:G42".

        For a one dimensional (only horizontal or only vertical) range of 
        cells, 'data' is a list. For a two dimensional range of cells, 'data' 
        is a list within a list. For example setting the 'cell_ref' "A1:C3"
        requires 'data' of the format:
        [[A1, B1, C1], [A2, B2, C2], [A3, B3, C3]].
        """

        self._check_for_lock()
        
        r = self._cell_range_to_index(cell_ref)
        sheet = self.spreadsheet.sheets[sheet]

        if r["row_start"] == r["row_end"]: # A row of cells
            self._check_1D_list(data)
            sheet[r["row_start"],
                  r["column_start"]:r["column_end"] + 1].values = data
        
        elif r["column_start"] == r["column_end"]: # A column of cells
            self._check_1D_list(data)
            sheet[r["row_start"]:r["row_end"] + 1,
                  r["column_start"]].values = data
        
        else: # A grid of cells
            self._check_list(data)
            for x, row in enumerate(data):
                if not isinstance(row, list):
                    raise ValueError("Expected a list of cells.")
                
                for y, cell in enumerate(row):
                    data[x][y] = self._check_numeric(cell)
                
            sheet[r["row_start"]:r["row_end"]+1,
                  r["column_start"]:r["column_end"]+1].values = data

            
    def get_cells(self, sheet, cell_ref):
        """Gets the value(s) of a single cell or a cell range. This can be used
        when it is not known if 'cell_ref' refers to a single cell or a range.

        See 'get_cell' and 'get_cell_range' for more information.
        """

        if self._is_single_cell(cell_ref):
            return self.get_cell(sheet, cell_ref)
        else:
            return self.get_cell_range(sheet, cell_ref)

            
    def get_cell(self, sheet, cell_ref):
        """Returns the value of a single cell.

        'cell_ref' is what one would use in LibreOffice Calc. Eg. "ABC945".
        """

        self._check_single_cell(cell_ref)
        
        r = self._cell_to_index(cell_ref)
        sheet = self.spreadsheet.sheets[sheet]
        
        return sheet[r["row_index"], r["column_index"]].value
            

    def get_cell_range(self, sheet, cell_ref):
        """Returns the values of a range of cells.
        
        'cell_ref' is what one would use in LibreOffice Calc. Eg. "ABC945".
        """

        r = self._cell_range_to_index(cell_ref)
        sheet = self.spreadsheet.sheets[sheet]

        # Cell ranges are requested as: [vertical area, horizontal area]
        
        if r["row_start"] == r["row_end"]: # A row of cells was requested
            return sheet[r["row_start"],
                         r["column_start"]:r["column_end"] + 1].values

        elif r["column_start"] == r["column_end"]: # A column of cells
            return sheet[r["row_start"]:r["row_end"]+1,r["column_start"]].values
        
        else: # A grid of cells
            return sheet[r["row_start"]:r["row_end"] + 1,
                         r["column_start"]:r["column_end"] + 1].values
        

        
    def save_spreadsheet(self, filename, directory="./saved_spreadsheets/"):
        """Save the spreadsheet in it's current state.
        
        'filename' is the name of the file.
        
        """

        if self.lock.locked():
            self.spreadsheet.save(directory + filename)
        else:
            return "Spreadsheet not locked"
