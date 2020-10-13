# -*- coding: utf-8 -*-
#pip install python-docx icalendar

from datetime import datetime, timedelta
from docx import Document
from icalendar import Calendar, Event
import os
import sys
import pandas as pd
import re
import pprint


year_now = datetime.now().strftime("%Y")
mounth_now = datetime.now().strftime("%m")


path = os.getcwd()
try:
    os.makedirs(f'{path}\\render\\', exist_ok=True)
except:
    pass


try:
    os.makedirs(f'{path}\\put docs here\\', exist_ok=True)
except:
    pass
render = '\\put docs here\\'


dir_list = [os.path.join(path + render, x) for x in os.listdir(path + render)]
docx_file = [el.split('\\') for el in dir_list]

for i in range(len(docx_file)):
    docx_name = docx_file[i][-1]

    document = Document(f'{path}{render}{docx_name}')

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

        elif e_date == ''.join(pattern_3):
            e_date = e_date.split()
            e_date_day = e_date[0]
            mounth_dict = {'01': 'января', '02': 'февраля', '03': 'марта', '04': 'апреля', '05': 'мая', '06': 'июня', '07': 'июля',
                           '08': 'августа', '09': 'сентября', '10': 'октября', '11': 'ноября', '12': 'декабря'}
            for k,v in mounth_dict.items():
                if v == e_date[1]:
                    e_date_mounth = k
            e_date_year = year_now
            # print(e_date_day, e_date_mounth)
            # print(e_date)

        e_from = row_data['From'].strip().replace('\\n', '')
        if (int(year_now) - int(e_date_year)) % 5 == 0 and (int(year_now) - int(e_date_year)) >= 50 and int(e_date_year) != int(year_now):
            e_from = f'ЮБИЛЕЙ {int(year_now) - int(e_date_year)} ЛЕТ! ({e_date_year} г.р.) {e_from}'
        elif int(e_date_year) == int(year_now):
            e_from = f'(Г.р. не известен) {e_from}'
        else:
            e_from = f'({e_date_year} г.р.) {e_from}'


        # e_name = row_data['Name'].strip().replace('  ', ' ')
        e_name = row_data['Name'].strip()

        pprint.pprint(e_name)
        row_data['summary'] = f'ДР {e_name}. {e_from}'
        row_data['description'] = ''

        start = datetime(int(year_now), int(e_date_mounth), int(e_date_day))
        end = datetime(int(year_now), int(e_date_mounth), int(e_date_day))
        end = end + timedelta(days=1)
        print(start)
        print(end)

        row_data[u'dtstart'] = start
        row_data[u'dtend'] = end

        del row_data['Name']
        del row_data['Date']
        del row_data['From']

        data.append(row_data)
   # pprint.pprint(data)


    cal = Calendar()

    for row in data:
        event = Event()

        event.add('summary', row['summary'])
        event.add('dtstart', row['dtstart'])
        event.add('dtend', row['dtend'])
        event.add('description', row['description'])
        # event.add('location', row['Room'])
        event.add('rrule', {'freq': 'yearly'})
        cal.add_component(event)




    ical_name = docx_name.split('.')[0]
    f = open(f'{path}\\render\\{ical_name}.ics', 'wb')
    f.write(cal.to_ical())
    f.close()