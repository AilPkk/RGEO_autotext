# This script reads text from statistics and creates paragraphs of text
import openpyxl as opx
import csv

xls_name = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Geoloogiline lõige koos statistikaga.xlsx"
dataframe = []  # declare empty df


# reads data

workbook = opx.load_workbook(xls_name, read_only=True, data_only=True)
sheet = workbook.active

for row in sheet.values:
    dataframe.append(list(row))

workbook.close()

# cleanup
dataframe_filtered = [x for x in dataframe if x[3] is not None or x[1] is not None]
for line in dataframe_filtered:
    while line[-1] is None:
        del line[-1]


# Data will be sorted into respective tables

split_index = [0]
table_headers = ["Teepinnast sügavus (m)", "Abs (m)", "Kaetud kihi nr", "Kiht esines puuraukudes "]

for i in range (len(dataframe_filtered)-1):
    try:
        if dataframe_filtered[i][3] in table_headers:
            split_index.append(i)
    except: pass

paksus_m = dataframe_filtered[:split_index[1]]
sygavus_m = dataframe_filtered[split_index[1]:split_index[2]]
abs_m =  dataframe_filtered[split_index[2]:split_index[3]]
kaetud_nr =  dataframe_filtered[split_index[3]:split_index[4]]
esines_nr =  dataframe_filtered[split_index[4]:]

#    print(len(paksus_m), len(sygavus_m), len(abs_m), len(kaetud_nr), len(esines_nr))
#    print(kaetud_nr)

def write_text(table):
    # Creates file and fills it with necessary info
    pass


# debug sh1t
# csv_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\test.csv"
# with open(csv_path, "w", newline="") as f:
#     write = csv.writer(f)
#     write.writerows(dataframe_filtered)

