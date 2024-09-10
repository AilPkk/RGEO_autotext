# This script reads text from statistics and creates paragraphs of text
# TODO: make it pretty again
# TODO: cases of nouns
# TODO: Preserve layer types
from time import process_time_ns

import openpyxl as opx
from pathlib import WindowsPath
from string import digits

# xls_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Geoloogiline lõige koos statistikaga.xlsx"
# kihid_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\sample\\Kihtide alus.txt"

dataframe = []  # declare empty df

# Ask for folder and generate full file paths
#workfolder_path = input("Sisesta töökausta asukoht (shift+parem klahv -> Copy as path): ")
workfolder_path = "C:\\Users\\saara\\OneDrive\\Töölaud\\Rakendusgeoloogia\\Ailar\\py\\autotext\\samplee"

xls_path = workfolder_path[0:-1]+"\\Geoloogiline lõige koos statistikaga.xlsx\""
kihid_path = workfolder_path[0:-1]+"\\Kihtide alus.txt\""
out_path = workfolder_path[0:-1]+"\\tulem.txt\""
xls_path = WindowsPath(xls_path.replace('"', ''))
kihid_path = WindowsPath(kihid_path.replace('"', ''))
out_path = WindowsPath(out_path.replace('"', ''))


### read data

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


## sort data

split_index = [0]
table_headers = ["Teepinnast sügavus (m)", "Abs (m)", "Kaetud kihi nr", "Kiht esines puuraukudes "]

for i in range (len(dataframe_filtered)-1):
    try:
        if dataframe_filtered[i][3] in table_headers:
            split_index.append(i)
    except: pass

UP_list = dataframe_filtered[split_index[1]][4:-3]
paksus_m = dataframe_filtered[1:split_index[1]-3]
sygavus_m = dataframe_filtered[split_index[1]+1:split_index[2]]
abs_m =  dataframe_filtered[split_index[2]+1:split_index[3]]

kaetud_nr =  dataframe_filtered[split_index[3]+1:split_index[4]]
kaetud_loetelu = [str(item[-1]).split(",") for item in kaetud_nr]
for loetelu in kaetud_loetelu:
    try:
        loetelu.remove("0")
    except: pass

esines_nr = dataframe_filtered[split_index[4]+1:]
esines_loetelu = [(str(item[-1])).split(",") for item in esines_nr]

# make lists to index
layer_list = []
for row in abs_m:
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

kihtide_loetelu = [(str(item[0])) for item in esines_nr]

pindmine_kiht = [[] for _ in range(len(layer_list))]
for i in range(len(pindmine_kiht)-1): #kihid
    for j in range(4,len(esines_nr[1])-1):
        if esines_nr[i][j] is not None:
            appending = UP_list[j - 4]
            appending = ''.join(c for c in appending if c in digits)
            if not any(appending in sublist for sublist in pindmine_kiht):
                pindmine_kiht[i].append(appending)

for i in range(len(pindmine_kiht)):
    for kiht in pindmine_kiht[i]:
        if kiht in esines_loetelu[i]:
            esines_loetelu[i].remove(kiht)


### text gen

text_list = [[] for _ in range(len(layer_list))]

for i in range(len(layer_list)):
    text_list[i].append("KIHT %s: " % (", ".join(kihi_kirjeldus[i])))
    if len(pindmine_kiht[i]) > 0:
        text_list[i].append("Uuringualal on kiht pindmiseks kihiks uuringupunktide %s alal " % (", ".join(pindmine_kiht[i])))

    paksus_min = paksus_m[i][-2]
    paksus_max = paksus_m[i][-1]
    if paksus_min == paksus_max:
        text_list[i].append("%s paksuse kihina. " %(str(paksus_max)))
    else:
        text_list[i].append("%s kuni %s m paksuse kihina. " % (str(paksus_min), str(paksus_max)))

    if len(esines_loetelu[i]) > 0:
        text_list[i].append("Kiht avati uuringupunktide %s alal kihtide " % (", ".join(esines_loetelu[i])))
        appstr = []
        for kiht in kaetud_loetelu[i]:
            kirjeldus_index = kihtide_loetelu.index(kiht)
            appstr.append("%s (%s)" % (kiht, kihi_kirjeldus[kirjeldus_index][1]))
        appstr = "%s all " % (", ".join(appstr))
        text_list[i].append(appstr)

    sygavus_min = sygavus_m[i][-2]
    sygavus_max = sygavus_m[i][-1]
    abs_min = abs_m[i][-2]
    abs_max = abs_m[i][-1]

    if sygavus_max > 0:
        if sygavus_min == sygavus_max:
            text_list[i].append("Kiht lasub maapinnast %s m sügavusel " % (str(sygavus_max)))
        else:
            text_list[i].append("Kiht lasub maapinnast %s kuni %s m sügavusel, " % (str(sygavus_min), str(sygavus_max)))
        if abs_min == abs_max:
            text_list[i].append("abs. kõrgusel %s m." % (str(abs_max)))
        else:
            text_list[i].append("abs. kõrgusel %s kuni %s m." % (str(abs_min), str(abs_max)))

text_list.insert(0, ("1. SELETUSKIRI\n"+ #Lisab päise
                     "1.1. ÜLDOSA\n\n"+
                     "[ASUKOHT] teostati geoloogilised väliuuringud [KUUPÄEV]. Uuringud teostati mahus ja asukohas, mis olid Tellija poolt ette antud.\n\n"+
                     "Välitööde käigus puuriti puurmasinaga [PUURMASIN] %s puurauku, sügavusega [SÜGAVUS].\n\n" % (str(len(UP_list))) +
                     "Uuringupunktid mõõdeti välja mõõdulindiga alusplaanil märgitud objektidest või märgiti maha GPS seadmega Garmin Montana 750i.\n\n"+
                     "Kõrgusesse seoti uuringupunktid geoalusel näidatud kõrguspunktidest, kõrgused on EH2000 kõrgussüsteemis. Uuringupunktide koordinaadid, mis on võetud alusplaanilt, on L-EST97 koordinaatsüsteemis.\n\n"+
                     "Töö on koostatud vastavalt EVS-EN ISO 14688-1:2018 Eesti Standardile „Geotehniline uurimine ja katsetamine. Osa 1: „Identifitseerimine ja kirjeldamine“ ning osa 2: „Liigituspõhimõtted“. Pinnastele on antud hinnangulised kategooriad A…D.\n\n"+
                     "Uuringupunktide asukohad on näidatud asendiplaanidel, joonistel [JOONISED]. Pinnaste täpsed kirjeldused ja lasuvuspilt on toodud geoloogilistes tulpades ja joonistel [TULBAD].\n\n"+
                     "1.2.GEOLOOGILINE EHITUS\n\n"+
                     "Maastikulise liigituse järgi jääb uuringuala [PIIRKOND] piirkonda. Maapinna kõrgused jäid puuraukude suudmetel abs. kõrguste [KÕRGUSED] vahemikku.\n\n"+
                     "Tee mulle ja täitepinnased (tIV):\n"
))
text_list.append("1.3. EHITUSGEOLOOGILISED TINGIMUSED")

printlist = []
for ln in text_list:
    printlist.append("".join(ln))

with open(out_path, "w", encoding="UTF-8") as output:
    output.write("\n\n".join(printlist))


