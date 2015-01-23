import random
from spreadsheet_client import SpreadsheetClient

if __name__ == "__main__":
    testing_doc = "20141209 HQA testsheet 1 (GS) BM.xlsx"
    #testing_doc = "range_test.xlsx"
    client = SpreadsheetClient("localhost", 5555, testing_doc)
    for x in range(0, 1):
        
        options = client.get_cells('Input', "B4:B94")
        rand_opt = random.choice(options)
        print(rand_opt)
        client.set_cells("Input", "B2", rand_opt)
        print(client.get_cells("Calcsheet", "K7:K21"))
    client.disconnect()
