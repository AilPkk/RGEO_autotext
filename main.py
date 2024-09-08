# This script reads text from statistics and creates paragraphs of text
import openpyxl as opx
import csv


def read_xlsx(file):
    # reads and gives data coherent structure

    workbook = opx.load_workbook(file, read_only=True, data_only=True)
    sheet = workbook.active

    dataf = []

    for row in sheet.values:
        dataf.append(row)

    workbook.close()

# debug sh1t
    csv_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\test.csv"
    with open(csv_path, "w", newline="") as f:
        write = csv.writer(f)
        write.writerows(dataf)




"""Key lines:
A: Puuraugu number			
D: Teepinnast sügavus (m)
D: Abs (m)
D: Kaetud kihi nr
D: Kiht esines puuraukudes 
"""

def write_text(table):
    # Creates file and fills it with necessary info
    pass

filename = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Geoloogiline lõige koos statistikaga.xlsx"

if __name__ == '__main__':
    read_xlsx(filename)
    # write_text
    pass

