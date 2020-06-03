#!/usr/bin/env python

import os.path, time
from openpyxl import Workbook

path = os.getcwd()
excel_file = input("Escriu el nom del excel: \n")
excel_file += ".xlsx"
filesheet = os.path.join(path, excel_file)
wb = Workbook()

sheet = wb.active
files = next(os.walk(path))[2]
line = 1
def excel_insert(data, line):

    sheet["A" + str(line)] = data[0]
    sheet["B" + str(line)] = data[1]
    sheet["C" + str(line)] = data[2]

excel_insert(("Arxiu", "Data de creació","Data de modificació",), line)

for file in files:
    if file.startswith("files_excel_list"):
        continue
    line += 1
    modified = time.ctime(os.path.getmtime(file))
    created = time.ctime(os.path.getctime(file))
    excel_insert((file, created, modified,), line)

wb.save(filesheet)

