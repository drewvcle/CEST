import openpyxl
# IN TERMINAL: type pip install openpyxl
wb = openpyxl.load_workbook('result.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

txt = open('latest.log', 'r')
deletetxt = open('delete.txt', mode = 'wt')
deleteCoords = open('deleteCoords.txt', mode = 'wt')


totalTxt = txt.readlines()


#Extracts all lines of strings containing the word "Coords"
coordsTxt = list(filter(lambda x: 'Coords' in x, totalTxt))
#Extracts excess words before the word Coords
for i in range(len(coordsTxt)):
    coordsTxt[i] = coordsTxt[i][coordsTxt[i].find("Coords") + 8:]

for txt in coordsTxt:
    deletetxt.write(txt)



coords = []
count = []
items = []

for i in range(len(coordsTxt)):
    coords.append(coordsTxt[i][:coordsTxt[i].find(":")])
    spec_items = []
    spec_count = []
    while coordsTxt[i] != '':      
        #spec_count.append(coordsTxt[i][coordsTxt[i].find("Count:") + 6 :coordsTxt[i].find("b,")])
        delim1b = coordsTxt[i].find("Count:") + 6
        coordsTxt[i] = coordsTxt[i][delim1b:]
        delim2b = coordsTxt[i].find("b,")
        if (delim1b == -1 or delim2b == -1):
            break
        spec_count.append(coordsTxt[i][:delim2b])  
        #items[i].append(coordsTxt[i][coordsTxt[i].find("minecraft:") + 10 :coordsTxt[i].find("\"}")])
        #coordsTxt[i] = coordsTxt[i][coordsTxt[i].find("}") + 1:]
        
        delim1 = coordsTxt[i].find("minecraft:") + 10
        coordsTxt[i] = coordsTxt[i][delim1:]
        delim2 = coordsTxt[i].find("\"}")
        if (delim1 == -1 or delim2 == -1):
            break
        spec_items.append(coordsTxt[i][:delim2])  
        

        # if (i == 0):
        #     print("REMAINDER --- " +  coordsTxt[i])

    items.append(spec_items)
    count.append(spec_count)
    

for txt in coords:
    deleteCoords.write(txt + '\n')


maxSize = [None] * 100

concat = []
for i in range(len(items)):
    spec_concat = []
    for j in range(len(items[i])):
        spec_concat.append("" + items[i][j])
    concat.append(spec_concat)


for i in range(len(concat)):
    for j in range(len(concat[i]), len(maxSize)):
        concat[i].append(None)

for i in range(len(coords), len(maxSize)):
    coords.append(None)


for i in range(26):
    sheet['F1'].value = "<-- Inner Chests"
    sheet['G1'].value = "Outer Chests -->"
    columnCell = str(chr(ord('A') + i))
    sheet[columnCell + str(2)].value = coords[i]
    #print('B' + str(1 + i))
    for j in range(len(concat)):
        sheet[columnCell + str(3 + j)].value = concat[i][j]

wb.save('result.xlsx')