import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
import os
import re 
from pprint import pprint



os.system('cls')
book = epub.read_epub('ref.epub')


fileText = []
scriptures = []

for item in book.get_items():
    if item.get_type() == ebooklib.ITEM_DOCUMENT:
        fileText.append(item.get_content())
        

file = open('sections.txt', 'w', encoding="utf-8")

for snippet in fileText:
    soup = BeautifulSoup(snippet, 'html.parser')
    
    for section in soup.find_all('div', {'class' : 'section'}):
        file.write(section.get_text())

    
    for test in soup.find_all('strong'):
        for section in test.find_all('a', href="#citation1"):
            test = re.sub("\xa0", " ", section.get_text())
            scriptures.append(test)


 
pprint(scriptures)

file.close()


file = open('sections.txt', 'r', encoding="utf-8")

sectionText = file.read()
x = re.sub("\u200b", " ", sectionText)
y = re.sub("\xa0", " ", x)

file.close()


file = open('sections.txt', 'w+')
file.write(y)
file.close()


file = open('sections.txt', 'r')
text = file.read()
x = re.findall("Song \d+", text)

songs = [] 
for song in x:
    songs.append(song)

pprint(songs)

talks = []
talkTime = {}

x = re.findall(".*\(\d+\smin.\).*\n|Song \d+", text)
for talk in x: 
    testRe = re.search("\(\d+\smin.\)", talk)

    removeRe = re.sub("\(\d+\smin.\)", "", talk)
    talks.append(removeRe)

    if testRe is not None:
        dotRemRe = re.sub("\.", "", testRe.group())
        talkTime[removeRe] = dotRemRe
    
    
file.close()

pprint(talkTime)


ministry = [[]]
treasures = [[]]
christians = [[]]

mIndex = 0
tIndex = 0
cIndex = 0

current = 'treasures'

for i in range(len(talks)):
    if current == 'christians' and 'Song' not in talks[i]:
        christians[cIndex].append(talks[i])
        if("Song" in talks[i + 1]):
            christians.append([])
            cIndex += 1
            current = 'treasures'
    if current == 'ministry' and 'Song' not in talks[i] and 'Concluding Comments' not in talks[i]:
        ministry[mIndex].append(talks[i])
        if("Song" in talks[i + 1]):
            ministry.append([])
            mIndex += 1
            current = 'christians'
    if current == 'treasures' and 'Song' not in talks[i] and 'Concluding Comments' not in talks[i]:
        treasures[tIndex].append(talks[i])
        if "Bible Reading:" in talks[i]:
            treasures.append([])
            tIndex += 1
            current = 'ministry'

ministry.pop()
treasures.pop()
christians.pop()

pprint(treasures)
print("==================")
pprint(ministry)
print("==================")
pprint(christians)

import openpyxl 
wb = openpyxl.load_workbook('template.xlsx')

ws1 = wb.active

songCounter = 0













print(ws1.max_row)
for row in range(1, ws1.max_row + 1):
    print(row)
    if(ws1[f'A{row}'].value == '[Song]'):
        ws1[f'A{row}'] = songs[songCounter]
        
        if(songCounter < len(songs) - 1):
            songCounter += 1
        

mIndex = 0
tIndex = 0
cIndex = 0

current = 'treasures'
rows = ws1.max_row
talkCell = {}

for row in range(1, 1000):
    if(ws1[f'C{row}'].value == '[Talk]'):
        if(current == 'treasures'):
            ws1[f'C{row}'] = ""
            for talk in reversed(treasures[tIndex]):
                rows += 1
                ws1.insert_rows(row, 1)
                ws1[f'C{row}'] = talk
                ws1[f'A{row}'] = talkTime[talk]   

                ws1[f'A{row}'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                ws1[f'C{row}'].alignment = openpyxl.styles.Alignment(wrap_text=True)

            if(tIndex < len(treasures) - 1):
                tIndex += 1
            current = 'ministry'
            continue

        if(current == 'ministry'):
            ws1[f'C{row}'] = ""
            for talk in reversed(ministry[mIndex]):
                rows += 1
                ws1.insert_rows(row, 1)
                ws1[f'C{row}'] = talk
                ws1[f'A{row}'] = talkTime[talk]

                ws1[f'A{row}'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                ws1[f'C{row}'].alignment = openpyxl.styles.Alignment(wrap_text=True)

            if(mIndex < len(ministry) - 1):
                mIndex += 1
            current = 'christians'
            continue

        if(current == 'christians'):
            ws1[f'C{row}'] = ""
            for talk in reversed(christians[cIndex]):
                rows += 1
                ws1.insert_rows(row, 1)
                ws1[f'C{row}'] = talk
                ws1[f'A{row}'] = talkTime[talk]

                ws1[f'A{row}'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                ws1[f'C{row}'].alignment = openpyxl.styles.Alignment(wrap_text=True)
            
            if(cIndex < len(christians) - 1):
                cIndex += 1
            current = 'treasures'

from datetime import timedelta
finishTime = timedelta(0, (7*3600+6*60))



# talkCell[talkTime[talk]] = f'D{row}'

pprint(talkCell)


# range = ws1['A1' : f'A{rows}']
# for cell in range:
#     for x in cell:
#         if("min" in x.value):






# for talk in talks:
#     try:
#         print(talkCell[talk])
#     except:
#         continue

from openpyxl.styles import Border, Side

top = Side(border_style='thin', color='FFFFFF')
left = Side(border_style='thin', color='FFFFFF')
right = Side(border_style='thin', color='FFFFFF')
bottom = Side(border_style='thin', color='FFFFFF')
border = Border(top = top, bottom = bottom, left = left, right = right)

cellRange = ws1['A1' : f'D{rows}']
for cell in cellRange:
    for x in cell:
        x.border = border
    
# right = Side(border_style='thin', color='000000')
# border = Border(top = top, bottom = bottom, left = left, right = right)

# columns = ['A', 'B', 'C', 'D']

# for i in columns:
#     cellRange = ws1[f'{i}3' : f'{i}{rows}']
#     for cell in cellRange:
#         for x in cell:
#             x.border = border

bottom = Side(border_style='thin', color='808080')
border = Border(top = top, bottom = bottom, left = left, right = right)


for row in range(1, 200):
    try:
        if('min' in ws1[f'A{row}'].value):
            stripNum = re.search('\d+', ws1[f'A{row}'].value)
            timeAdd = timedelta(0, (0*3600+int(stripNum.group())*60))
            finishTime += timeAdd
            ws1[f'D{row}'] = finishTime
            
            number_format = '[hh]:mm'
            ws1[f'D{row}'].number_format = number_format
            ws1[f'D{row}'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

            for key in talkTime:
                if(talkTime[key] == ws1[f'A{row}'].value):
                    finishTime = timedelta(0, (7*3600+6*60))

            cellRange = ws1[f'A{row}' : f'D{row}']
            for cell in cellRange:
                for x in cell:
                    x.border = border
    except:
        continue


first = re.sub("-.*\s", " ", book.get_metadata('DC', 'title')[0][0])
second = re.sub(",\s.*-", ", ", book.get_metadata('DC', 'title')[0][0])


ws1['B1'] = first
print(rows)

scripturesIndex = 0

for row in range(2, rows):
    if(ws1[f'B{row}'].value == '[Heading]'):
        print('here')
        ws1[f'B{row}'] = second
    
    if(ws1[f'C{row}'].value == '[Scripture]'):
        ws1[f'C{row}'] = scriptures[scripturesIndex]
        scripturesIndex += 1



wb.save(filename = "new.xlsx")

os.startfile('new.xlsx')


