#!/usr/bin/env python3

from datetime import datetime
from html.parser import HTMLParser
from openpyxl import Workbook
import os

now = datetime.now()
current_time = now.strftime("%Y%m%d")

#with open("Logistics_133800_20200528.xls") as f: // use for testing

# add path to file
path = r'C:\Users\Username\Path\To\Logistics_133800_' + current_time + '.xls'
with open(path) as f: 
    data = f.read()

class CoopHTMLParser(HTMLParser):
    parsed_data = ''
    start_tag = ''
    is_t_head = False
    is_t_body = False
    is_span = False
    firstRow = []
    count = 0
    orders_list = []
    order = {}
    
    def handle_starttag(self, tag, attrs):
        if tag == 'span':
            self.is_span = True
        if tag == 'thead':
            self.is_t_head = True
            self.is_t_body = False
        if tag == 'tbody':
            self.is_t_body = True
            self.is_t_head = False
        start_tag = tag

    def handle_data(self, data):
        if not self.is_span:
            data = data.strip()
            if self.count == 1:
                start_time, sep, end_time = data.partition('-')
                start_time = start_time.strip() 
                end_time = end_time.strip()
                data = start_time + sep + end_time
            if (self.parsed_data == 'Avhämtning' or self.parsed_data == 'Hemleverans') and data == '':    
                pass
            else:
                self.parsed_data = data

    def handle_endtag(self, tag):
        if tag == 'span':
            self.is_span = False
        if self.is_t_head and tag == 'th':
            self.firstRow.append(self.parsed_data)
        if self.is_t_body and tag == 'td':
            if self.count < 18:
                self.order.update({ self.firstRow[self.count] : self.parsed_data })
                self.count += 1
                if self.count == 11:
                    self.parsed_data = ''
                if self.count == 18:
                    self.orders_list.append(self.order.copy())
                    self.count = 0

p = CoopHTMLParser()
try:
    p.feed(data)
except TypeError:
    print()

kommunalwb = Workbook()
kommunalws = kommunalwb.active
avhamtningwb = Workbook()
avhamtningws = avhamtningwb.active
kommunal = []
avhamtning = []

for x in range(0, len(p.orders_list) - 1):
    if p.orders_list[x]['Leveranstyp'] == 'Avhämtning':
        avhamtning.append(p.orders_list[x])
    else:
        kommunal.append(p.orders_list[x])

def switch_column(argument):
    switcher = {
        'A': '#',
        'B': 'Lev tid',
        'C': 'Exakt lev tid',
        'D': 'Id',
        'E': 'Kundtyp',
        'F': 'Kundnamn',
        'G': 'Gata',
        'H': 'Postnr',
        'I': 'Ort',
        'J': 'Rutt',
        'K': 'Leveranstyp',
        'L': 'Enhet',
        'M': 'Beställare' ,
        'N': 'Speditör kommentar',
        'O': 'Orderkommentar',
        'P': 'Kund Telefon',
        'Q': 'Enhet Telefon',
        'R': 'Ordersumma (inkl moms)',
    }
    return switcher.get(argument, 'felfelfel')

# add separation between each 'time-slot'.
def addSeparation(list):
    for i in range(0, len(list) -1):
        if i != 0:
            if not list[i - 1]['Lev tid'] == '':
                if list[i]['Lev tid'].lower() != list[i - 1]['Lev tid'].lower():
                    list.insert(i, {'#': '','Lev tid': '','Exakt lev tid': '','Id': '', 'Kundtyp': '', 'Kundnamn': '', 'Gata': '', 'Postnr': '',
                    'Ort': '', 'Rutt': '', 'Leveranstyp': '', 'Enhet': '', 'Beställare': '', 'Speditör kommentar': '', 'Orderkommentar': '', 'Kund Telefon': '','Enhet Telefon': '', 'Ordersumma (inkl moms)': ''})

# add the first row describing each coloumn
def add_first_row(from_list, to_work_sheet,):
    for i in range(0, len(from_list) - 1):
        to_work_sheet[str(chr(i + 65)) + str(1)] = from_list[i]

# add valuse to columns 
def addToColumn(list, work_sheet):
    for i in range(0, len(list) - 1):
        for y in range(65, len(list[i]) + 64):
            work_sheet[str(chr(y)) + str(i + 2)] = list[i][switch_column(str(chr(y)))]

kommunal = sorted(kommunal, key=lambda i: i['Orderkommentar'].lower())
avhamtning = sorted(avhamtning, key=lambda i: i['Lev tid'])
addSeparation(avhamtning)
add_first_row(p.firstRow, kommunalws)
add_first_row(p.firstRow, avhamtningws)
addToColumn(avhamtning, avhamtningws)
addToColumn(kommunal, kommunalws)

kommunal_path = r'G:\' #Add path to folders
avhamtning_path = r'G:\' #Add path to folders
kommunal_file_name = "Kommunal." + current_time + ".xlsx"
avhamtning_file_name = "Avhämtning." + current_time + ".xlsx"

kommunalwb.save(kommunal_path + kommunal_file_name)
avhamtningwb.save(avhamtning_path + avhamtning_file_name)
os.system('start "excel" "{kommunal_path}{kommunal_file_name}"')
os.system('start "excel" "{avhamting_path}{avhamtning_file_name}"')