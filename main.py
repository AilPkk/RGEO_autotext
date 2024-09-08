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
    # Data will be sorted into respective tables

    split_index = [0]
    table_headers = ["Teepinnast sügavus (m)", "Abs (m)", "Kaetud kihi nr", "Kiht esines puuraukudes "]

    for i in range (len(dataf_filtered)-1):
        try:
            if dataf_filtered[i][3] in table_headers:
                split_index.append(i)
        except: pass

    paksus_m = dataf_filtered[:split_index[1]]
    sygavus_m = dataf_filtered[split_index[1]:split_index[2]]
    abs_m =  dataf_filtered[split_index[2]:split_index[3]]
    kaetud_nr =  dataf_filtered[split_index[3]:split_index[4]]
    esines_nr =  dataf_filtered[split_index[4]:]


def write_text(table):
    # Creates file and fills it with necessary info
    pass

filename = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Geoloogiline lõige koos statistikaga.xlsx"
dataframe = []  # declare empty df

if __name__ == '__main__':
    dataframe_filtered = read_xlsx(filename, dataframe)
    organize_data(dataframe_filtered)
    # write_text
    pass

# debug sh1t
# csv_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\test.csv"
# with open(csv_path, "w", newline="") as f:
#     write = csv.writer(f)
#     write.writerows(dataframe_filtered)

