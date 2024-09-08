# This script reads text from statistics and creates paragraphs of text
import openpyxl as opx


def read_xlsx(file):
    # reads and gives data coherent structure

    workbook = opx.load_workbook(file, read_only=True, data_only=True)
    sheet = workbook.active

    max_row = sheet.max_row
    max_col = sheet.max_column

    print(max_row, max_col)

    dataf = [[] for _ in range(max_col + 1)]

    for i in range (1, max_row + 1):
        print(round(i / max_row * 100, 1))
        if sheet.cell(row = i, column = 4).value is not None:

            for j in range (1, max_col + 1):
                try:
                    dataf[i].append(sheet.cell(row = i, column = j).value)
                except:
                    pass

    print(dataf)



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
