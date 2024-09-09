# This script reads text from statistics and creates paragraphs of text
import openpyxl as opx
import csv
from pathlib import WindowsPath

# xls_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Geoloogiline lõige koos statistikaga.xlsx"
# kihid_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Kihtide alus.txt"

dataframe = []  # declare empty df

# Ask for folder and generate full file paths
workfolder_path = input("Sisesta töökausta asukoht (shift+parem klahv -> Copy as path): ")
print(workfolder_path)
xls_path = workfolder_path[0:-1]+"\\Geoloogiline lõige koos statistikaga.xlsx"+"\""
kihid_path = workfolder_path[0:-1]+"\\Kihtide alus.txt"+"\""
xls_path = WindowsPath(xls_path.replace('"', ''))
kihid_path = WindowsPath(kihid_path.replace('"', ''))

print(xls_path)
print(kihid_path)

# reads data

workbook = opx.load_workbook(xls_path, read_only=True, data_only=True)
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

# make lists to index
UP_list = kaetud_nr[0][4:-1] # just in case
layer_list = []
for row in kaetud_nr:
    if row[0] is not None:
        layer_list.append(row[0])

# get layer description
kihi_kirjeldus = []
with open(kihid_path, "r", encoding="UTF-8") as kihid:
    for kiht in kihid:
        kiht = kiht.strip()
        kiht = kiht.rsplit(". ")
        if len(kiht) > 1:
            kihi_kirjeldus.append(kiht)



# text gen


"""
KIHT 6, Kruusane (mölline) eriteraline LIIV (gr(si)Sa, fglIII): 
PA-1...-2, PA-8 ja PA-10 alal avati täitepinnase (kiht 3) või vähese orgaanilise aine sisaldusega 
liivase MÖLLI (kiht 5) all, teepinnast 0,40...0,45 meetri sügavusel 0,20...1,10 meetri paksune kruusase (möllise) 
eriteralise liiva kiht, abs. kõrgusel 62,63...66,10 meetrit. Kiht on helepruun, kesktihe kuni tihe, 
niiske, sisaldab jämepurdu 15...20%. 
Kiht on mõõdukalt külmaohtlik ning ei täida dreenimistingimusi.
"""


def write_text(table):
    # Creates file and fills it with necessary info
    pass


# debug sh1t
# csv_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\test.csv"
# with open(csv_path, "w", newline="") as f:
#     write = csv.writer(f)
#     write.writerows(dataframe_filtered)

