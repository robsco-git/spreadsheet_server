import logging
from math import pow
from com.sun.star.uno import RuntimeException

class SpreadsheetConnection:
    """Handles connections to the workbooks opened by soffice."""
    def __init__(self, workbook, lock):
        self.workbook = workbook
        self.lock = lock

    # The lock/unlock functions must be used for the getting and setting functions to work
    # In order for simultaneous requests to the same workbook to not interfere with each other
    def lock_workbook(self):
        self.lock.acquire()

    def _range_to_index(self, cell_range):
        """ Used to convert a cell (or range) reference to indexes """
        def split_char_num(cell_ref):
            num = ""
            str_val = 0
            split = False
            for x, char in enumerate(cell_ref):
                if char.isdigit():
                    if not split:
                        split = True
                        # calculate index from string part of refernece
                        string_part = cell_ref[:x]
                        for i, c in enumerate(string_part):
                            if i == len(string_part) - 1:
                                # least significant character
                                str_val += ord(c.upper()) - 65
                            else:
                                str_val += (ord(c.upper()) - 65 + 1) * pow(26, len(string_part)-i-1)
                    # handle numeric part of the cell reference            
                    num += char
            return str_val, num

        if len(cell_range.split(':')) == 1:
            # working with a single cell reference
            str_val, num = split_char_num(cell_range)
            return [int(num)-1, str_val]
        else:
            left, right = cell_range.split(':')
            left_str_val, left_num = split_char_num(left)
            right_str_val, right_num = split_char_num(right)
            return [[int(left_num)-1, int(right_num)-1],[left_str_val, right_str_val]]

    def set_cells(self, sheet, cell_range, data):
        if self.lock.locked():
            r = self._range_to_index(cell_range)

            # check if each cell is a number of not - try catch float()
            
            sheet = self.workbook.sheets[sheet]
            try:
                if len(cell_range.split(':')) == 1:
                    # a single cell
                    try:
                        data = float(data)
                    except ValueError:
                        pass
                    sheet[r[0],r[1]].value = data
                else:
                    if r[0][0] == r[0][1]: # A row of cells is being set
                        for x, cell in enumerate(data):
                            try:
                                data[x] = float(cell)
                            except ValueError:
                                pass
                        sheet[r[0][0],r[1][0]:r[1][1]+1].values = data
                    elif r[1][0] == r[1][1]: # A column of cells is being set
                        for x, cell in enumerate(data):
                            try:
                                data[x] = float(cell)
                            except ValueError:
                                pass
                        sheet[r[0][0]:r[0][1]+1,r[1][0]].values = data
                    else: # A grid of cells is being set
                        for x, lst in enumerate(data):
                            for y, cell in enumerate(data):
                                try:
                                    data[x][y] = float(cell)
                                except ValueError:
                                    pass
                        sheet[r[0][0]:r[0][1]+1,r[1][0]:r[1][1]+1].values = data
                return True
            except RuntimeException:
                return False
        else:
            return False
        
    def get_cells(self, sheet, cell_range):
        # Workbook does not need to be locked cells to be read
        try:
            r = self._range_to_index(cell_range)
            sheet = self.workbook.sheets[sheet]
            # [vertical area, horizontal area]
            if len(cell_range.split(':')) == 1:
                # a single cell
                return sheet[r[0],r[1]].value
            else:
                if r[0][0] == r[0][1]: # A row of cells was requested
                    return sheet[r[0][0],r[1][0]:r[1][1]+1].values
                elif r[1][0] == r[1][1]: # A column of cells was requested
                    return sheet[r[0][0]:r[0][1]+1,r[1][0]].values
                else: # A grid of cells was requested
                    return sheet[r[0][0]:r[0][1]+1,r[1][0]:r[1][1]+1].values
        except:
            return False

    def unlock_workbook(self):
        try:
            self.lock.release()
            return True
        except RuntimeError:
            return False
        
    def save_workbook(self, filename):
        if self.lock.locked():
            self.workbook.save("./saved_workbooks/" + filename)
        else:
            return "Workbook not locked"
