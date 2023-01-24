import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
import os
import re 
from pprint import pprint



os.system('cls')
book = epub.read_epub('ref.epub')


fileText = []


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
            print(section.get_text())
        




            
# if(subSection.get_text() == '<a epub:type="noteref" href="#citation1">'):