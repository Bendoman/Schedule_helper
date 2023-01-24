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


from tkinter import *
from tkinter import ttk

# if you are still working under a Python 2 version, 
# comment out the previous line and uncomment the following line
# import Tkinter as tk

root = Tk()
root.title('Hello world')
root.geometry('1000x500')

# Create a main frame
main_frame = Frame(root)
main_frame.pack(fill=BOTH, expand=1)

# Create a canvase
my_canvas = Canvas(main_frame)
my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

# Add a scrollbar to the Canvas
my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
my_scrollbar.pack(side=RIGHT, fill=Y)

# configure the canvas
my_canvas.configure(yscrollcommand=my_scrollbar.set)
my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox('all')))

# create another frame inside the canvas
second_frame = Frame(my_canvas)

my_canvas.create_window((0,0), window=second_frame, anchor="nw")



for i in range(0, len(talks)):
    try:
        test = re.search(".{95}", talks[i]).group()
        test += "..."
    except:
        test = talks[i]


    Label(second_frame, text=test).grid(row=i, column=0, pady=10, padx=10)
    Button(second_frame, text=f'Button {i}').grid(row=i, column=1, pady=10, padx=10)


def printInput():
    inp = root.get(1.0, "end-1c")
    print(inp)

test1 = Text(second_frame, height=5, width=20)
test1.grid(row=0, column=2)

def printInput():
    print(test1.get(1.0, "end-1c"))

button = Button(second_frame, text="Print", command=printInput)
button.grid(column=2, row=1)

# menu= StringVar()
# menu.set("Select Any Language")

# #Create a dropdown Menu
# drop= OptionMenu(second_frame, menu,"C++", "Java","Python","JavaScript","Rust","GoLang")
# drop.grid(column=3, row=0)


# variable = StringVar(root)
# variable.set("one") # default value

# w = ttk.Combobox(second_frame, textvariable=variable, values=["Carrier 19EX 4667kW/6.16COP/Vanes", "Carrier 19EX 4997kW/6.40COP/Vanes", "Carrier 19EX 5148kW/6.34COP/Vanes", "Carrier 19EX 5208kW/6.88COP/Vanes", "Carrier 19FA 5651kW/5.50COP/Vanes", "Carrier 19XL 1674kW/7.89COP/Vanes", "Carrier 19XL 1779kW/6.18COP/Vanes", "Carrier 19XL 1797kW/5.69COP/Vanes", "Carrier 19XL 1871kW/6.49COP/Vanes", "Carrier 19XL 2057kW/6.05COP/Vanes", "Carrier 19XR 1076kW/5.52COP/Vanes", "Carrier 19XR 1143kW/6.57COP/VSD", "Carrier 19XR 1157kW/5.62COP/VSD", "Carrier 19XR 1196kW/6.50COP/Vanes", "Carrier 19XR 1213kW/7.78COP/Vanes", "Carrier 19XR 1234kW/5.39COP/VSD", "Carrier 19XR 1259kW/6.26COP/Vanes", "Carrier 19XR 1284kW/6.20COP/Vanes", "Carrier 19XR 1294kW/7.61COP/Vanes", "Carrier 19XR 1350kW/7.90COP/VSD", "Carrier 19XR 1403kW/7.09COP/VSD", "Carrier 19XR 1407kW/6.04COP/VSD", "Carrier 19XR 1410kW/8.54COP/VSD", "Carrier 19XR 1558kW/5.81COP/VSD", "Carrier 19XR 1586kW/5.53COP/VSD", "Carrier 19XR 1635kW/6.36COP/Vanes", "Carrier 19XR 1656kW/8.24COP/VSD", "Carrier 19XR 1723kW/8.32COP/VSD", "Carrier 19XR 1727kW/9.04COP/Vanes", "Carrier 19XR 1758kW/5.86COP/VSD", "Carrier 19XR 1776kW/8.00COP/Vanes", "Carrier 19XR 1801kW/6.34COP/VSD", "Carrier 19XR 2391kW/6.44COP/VSD", "Carrier 19XR 2391kW/6.77COP/Vanes", "Carrier 19XR 742kW/5.42COP/VSD", "Carrier 19XR 823kW/6.28COP/Vanes", "Carrier 19XR 869kW/5.57COP/VSD", "Carrier 19XR 897kW/6.23COP/VSD", "Carrier 19XR 897kW/6.50COP/Vanes", "Carrier 19XR 897kW/7.23COP/VSD", "Carrier 23XL 1062kW/5.50COP/Valve", "Carrier 23XL 1108kW/6.92COP/Valve", "Carrier 23XL 1196kW/6.39COP/Valve", "Carrier 23XL 686kW/5.91COP/Valve", "Carrier 23XL 724kW/6.04COP/Vanes", "Carrier 23XL 830kW/6.97COP/Valve", "Carrier 23XL 862kW/6.11COP/Valve", "Carrier 23XL 862kW/6.84COP/Valve", "Carrier 23XL 865kW/6.05COP/Valve", "Carrier 30RB100 336.5kW/2.8COP", "Carrier 30RB110 371kW/2.8COP", "Carrier 30RB120 416.4kW/2.8COP", "Carrier 30RB130 447.7kW/2.8COP", "Carrier 30RB150 507.8kW/2.8COP", "Carrier 30RB160 538kW/2.9COP", "Carrier 30RB170 585.5kW/2.8COP", "Carrier 30RB190 662.9kW/2.8COP", "Carrier 30RB210 710kW/2.9COP", "Carrier 30RB225 753.3kW/2.8COP", "Carrier 30RB250 836.2kW/2.8COP", "Carrier 30RB275 915kW/2.8COP", "Carrier 30RB300 993.8kW/2.8COP", "Carrier 30RB315 1076.1kW/2.9COP", "Carrier 30RB330 1123.6kW/2.8COP", "Carrier 30RB345 1170.7kW/2.8COP", "Carrier 30RB360 1248.4kW/2.8COP", "Carrier 30RB390 1325.8kW/2.8COP", "Carrier 30RB90 303.8kW/2.9COP", "Carrier 30XA100 330.1kW/3.1COP", "Carrier 30XA110 359.9kW/3COP", "Carrier 30XA120 389kW/3COP", "Carrier 30XA140 466.7kW/3.1COP", "Carrier 30XA160 535.1kW/3.1COP", "Carrier 30XA180 601.9kW/3.1COP", "Carrier 30XA200 681.7kW/3.1COP", "Carrier 30XA220 743.7kW/3.1COP", "Carrier 30XA240 801.6kW/3COP", "Carrier 30XA260 881.7kW/3.1COP", "Carrier 30XA280 943.4kW/3.1COP", "Carrier 30XA300 1010.2kW/3.1COP", "Carrier 30XA325 1077.4kW/3.1COP", "Carrier 30XA350 1138.7kW/3COP", "Carrier 30XA400 1348kW/3COP", "Carrier 30XA450 1499.5kW/2.9COP", "Carrier 30XA500 1609.4kW/2.9COP", "Carrier 30XA80 265.5kW/2.9COP", "Carrier 30XA90 297.8kW/3.1COP", "DOE-2 Centrifugal/5.50COP", "DOE-2 Reciprocating/3.67COP", "McQuay AGZ010BS 34.5kW/2.67COP", "McQuay AGZ013BS 47.1kW/2.67COP", "McQuay AGZ017BS 54.5kW/2.67COP", "McQuay AGZ020BS 71kW/2.67COP", "McQuay AGZ025BS 78.1kW/2.67COP", "McQuay AGZ025D 96kW/2.81COP", "McQuay AGZ029BS 95.7kW/2.67COP", "McQuay AGZ030D 111.1kW/2.81COP", "McQuay AGZ034BS 117.1kW/2.61COP", "McQuay AGZ035D 122.7kW/2.93COP", "McQuay AGZ040D 133.3kW/2.96COP", "McQuay AGZ045D 149.8kW/3.02COP", "McQuay AGZ050D 169.2kW/2.96COP", "McQuay AGZ055D 181.5kW/2.93COP", "McQuay AGZ060D 197.3kW/2.87COP", "McQuay AGZ065D 204.3kW/3.02COP", "McQuay AGZ070D 225.4kW/2.84COP", "McQuay AGZ075D 257.1kW/2.93COP", "McQuay AGZ080D 285.2kW/2.87COP", "McQuay AGZ090D 313.7kW/2.87COP", "McQuay AGZ100D 351kW/2.81COP", "McQuay AGZ110D 373.1kW/2.87COP", "McQuay AGZ125D 411.8kW/2.87COP", "McQuay AGZ130D 455.8kW/2.81COP", "McQuay AGZ140D 479kW/2.99COP", "McQuay AGZ160D 539.1kW/2.93COP", "McQuay AGZ180D 605.6kW/2.81COP", "McQuay AGZ190D 633.4kW/2.96COP", "McQuay PEH 1030kW/8.58COP/Vanes", "McQuay PEH 1104kW/8.00COP/Vanes", "McQuay PEH 1231kW/6.18COP/Vanes", "McQuay PEH 1635kW/7.47COP/Vanes", "McQuay PEH 1895kW/6.42COP/Vanes", "McQuay PEH 1934kW/6.01COP/Vanes", "McQuay PEH 703kW/7.03COP/Vanes", "McQuay PEH 819kW/8.11COP/Vanes", "McQuay PFH 1407kW/6.60COP/Vanes", "McQuay PFH 2043kW/8.44COP/Vanes", "McQuay PFH 2124kW/6.03COP/Vanes", "McQuay PFH 2462kW/6.67COP/Vanes", "McQuay PFH 3165kW/6.48COP/Vanes", "McQuay PFH 4020kW/7.35COP/Vanes", "McQuay PFH 932kW/5.09COP/Vanes", "McQuay WDC 1973kW/6.28COP/Vanes", "McQuay WSC 1519kW/7.10COP/Vanes", "McQuay WSC 1751kW/6.73COP/Vanes", "McQuay WSC 471kW/5.89COP/Vanes", "McQuay WSC 816kW/6.74COP/Vanes", "Multistack MS 172kW/3.67COP/None", "Trane CGAM100 337.6kW/3.11COP", "Trane CGAM110 367.2kW/3.02COP"])
# w.grid(column=3, row=0)

lst = ['C', 'C++', 'Java',
       'Python', 'Perl',
       'PHP', 'ASP', 'JS']


def check_input(event):
    value = event.widget.get()

    if value == '':
        combo_box['values'] = lst
    else:
        data = []
        for item in lst:
            if value.lower() in item.lower():
                data.append(item)

        combo_box['values'] = data

# creating Combobox
combo_box = ttk.Combobox(second_frame)
combo_box['values'] = lst
combo_box.bind('<KeyRelease>', check_input)
combo_box.grid(column=3, row=0)


root.mainloop()



from datetime import timedelta
finishTime = timedelta(0, (7*3600+6*60))



# talkCell[talkTime[talk]] = f'D{row}'




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

            print("Does ", ws1[f'C{row}'], " require a speaker?")
            response = input()
            if(response != ""):
                ws1[f'B{row}'] = response




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

# os.startfile('new.xlsx')



