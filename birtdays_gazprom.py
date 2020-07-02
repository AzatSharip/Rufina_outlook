# -*- coding: utf-8 -*-
#pip install python-docx icalendar

from datetime import datetime
from docx import Document
from icalendar import Calendar, Event
import os
import pandas as pd



path = os.getcwd()
document = Document(f'{path}\\source_file.docx')
table = document.tables[0]

data = []
keys = {}

for i, row in enumerate(table.rows):
    text = (cell.text for cell in row.cells)

    if i == 0:
        keys = tuple(text)
        continue

    row_data = dict(zip(keys, text))
    print(row_data['Name'])