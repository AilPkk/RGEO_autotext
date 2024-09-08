# This script reads text from statistics and creates paragraphs of text
import openpyxl as opx
import csv


def read_xlsx(file, dataf):
    # reads data

    workbook = opx.load_workbook(file, read_only=True, data_only=True)
    sheet = workbook.active

    for row in sheet.values:
        dataf.append(list(row))

    # list cleanup
    dataf_filtered = [x for x in dataf if x[3] is not None or x[1] is not None]
    for line in dataf_filtered:
        while line[-1] is None:
            del line[-1]

    return dataf_filtered

def organize_data(dataf_filtered):
    pass

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
dataframe = []  # declare empty df

if __name__ == '__main__':
    dataframe_filtered = read_xlsx(filename, dataframe)
    # organize_data(dataframe_filtered)
    # write_text
    pass

# debug sh1t
csv_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\test.csv"
with open(csv_path, "w", newline="") as f:
    write = csv.writer(f)
    write.writerows(dataframe_filtered)

