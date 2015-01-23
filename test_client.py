import random
from spreadsheet_client import SpreadsheetClient

if __name__ == "__main__":
    testing_doc = "20141209 HQA testsheet 1 (GS) BM.xlsx"
    #testing_doc = "range_test.xlsx"
    try:
        client = SpreadsheetClient("localhost", 5555, testing_doc)
    except:
        print("Could not connect to server")
    else:
        try:
            for x in range(0, 100):
                options = client.get_cells('Input', "B4:B94")
                rand_opt = random.choice(options)
                print(rand_opt)
                client.set_cells("Input", "B2", rand_opt)
                cells = client.get_cells("Calcsheet", "K7:K21")
                for cell in cells:
                    print cell
        except:
            print("Error in connection")
        finally:
            client.disconnect()
