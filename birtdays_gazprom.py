# -*- coding: utf-8 -*-
#pip install python-docx icalendar

from datetime import datetime
from docx import Document
from icalendar import Calendar, Event
import os
import pandas as pd
import re


year_now = datetime.now().strftime("%Y")
mounth_now = datetime.now().strftime("%m")

path = os.getcwd()
document = Document(f'{path}\\source_file.docx')
table = document.tables[0]

data = []
keys = {}
c = 0
for i, row in enumerate(table.rows):
    c +=1
    text = (cell.text for cell in row.cells)

    if i == 0:
        keys = tuple(text)
        # print(keys)
        continue


    row_data = dict(zip(keys, text))



    e_date = row_data['Date'].strip()
    pattern_1 = re.findall(r'\d\d.\d\d.\d\d\d\d$', e_date)
    pattern_2 = re.findall(r'\d\d.\d\d.\d\d\d\d\s.*', e_date)
    pattern_3 = re.findall(r'\d*\s\w*', e_date)
    # print(e_date)

    if e_date == ''.join(pattern_1):
        e_date = e_date.split('.')
        e_date_day = e_date[0]
        e_date_mounth = e_date[1]
        e_date_year = e_date[2]

        # print(e_date)

    elif e_date == ''.join(pattern_2):
        e_date = e_date.split()
        e_date = [e.split('.') for e in e_date]

        e_date_day = e_date[0][0]
        e_date_mounth = e_date[0][1]
        e_date_year = e_date[0][2]

        # print(e_date)

    # elif e_date == ''.join(pattern_3):
    #     print(e_date)



    # print(datetime.now().strftime("%d-%m-%Y"))


    row_data[u'dtstart'] = datetime(int(year_now), int(e_date_mounth), int(e_date_day), 8, 0, 0)
    row_data[u'dtend'] = datetime(int(year_now), int(e_date_mounth), int(e_date_day), 20, 0, 0)
#



#     if len(e_time) > 1:
#
#         row_data[u'dtstart'] = datetime(2017, int(e_date[1]), int(e_date[0]),
#                                         int(e_time[0]), 0, 0)
#         row_data[u'dtend'] = datetime(2017, int(e_date[1]), int(e_date[0]),
#                                       int(e_time[1]), 0, 0)
#
#     else:
#         row_data[u'dtstart'] = datetime(2017, int(e_date[1]), int(e_date[0]),
#                                         17, 1, 0)
#         row_data[u'dtend'] = datetime(2017, int(e_date[1]), int(e_date[0]),
#                                       17, 0, 0)
#
#     data.append(row_data)
#
# print(data)

# cal = Calendar()

# for row in data:
#     event = Event()
#
#     event.add('summary', row['Activity'])
#     event.add('dtstart', row['dtstart'])
#     event.add('dtend', row['dtend'])
#     event.add('description', row['Activity'])
#     event.add('location', row['Room'])
#     event.add('rrule', {'freq': 'yearly'})
#     cal.add_component(event)
#
# f = open('course_schedule.ics', 'wb')
# f.write(cal.to_ical())
# f.close()