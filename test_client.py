import random
from client_python2 import SpreadsheetClient

if __name__ == "__main__":
    testing_doc = "test.ods"
    try:
        client = SpreadsheetClient("localhost", 5555, testing_doc)
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
