#!/usr/bin/env python3
""""
Author: Colin McAllister (https://github.com/offsetkeyz)

This will calculate perks based on the excel schedule.
"""

from datetime import timedelta
import datetime
import json
import random
from time import strptime
from tracemalloc import start
import pytz
from cogs import authentication, classes
import requests
from openpyxl import load_workbook
import datetime
from dateutil.tz import *

schedule_wb = load_workbook(authentication.get_schedule_location())
perks_workbook = load_workbook(authentication.get_perks_sheet_location())
perks_ws = perks_workbook['iSOC Perks']
start_date = ''
end_date = ''

all_perks_names = []
all_perks_cells = []
all_isoc_names = []

def get_start_date():
    input_date = input().strip()
    try:
        dt_start = datetime.datetime.strptime(str(input_date), '%d %b %Y')
        dt_start = dt_start.replace(hour=13)
    except ValueError as e:
        print ("Incorrect format with: " + input_date)
        print('Please try again (ex. 07 Jul 2033): ')
        input_date = get_start_date()
    return dt_start

#strips leading and trailing spaces
def strip_spaces(name_in):
    if type(name_in) is not str:
        return name_in
    if name_in != None:
        name_in_stripped = name_in.strip()
        return name_in_stripped 
    return name_in


def is_canadian(employee_in):
    if find_name_in_perks(employee_in.name) < 124:
        return True
    return False

# Add numbers into correct column. Function takes name, night hours, 12H day hours
    #search for name and return row number. (cell.row)
def find_name_in_perks(name_in) -> int:
    for i in range(len(all_perks_names)):
        if strip_spaces(name_in) == strip_spaces(all_perks_names[i]):
            return all_perks_cells[i].row
    return -1

def compare_names():
    for i in range(0,5,1):
        for cell in schedule_wb.worksheets[i]['A']:
            try:
                if len(cell.value) > 1 and '>' not in str(cell.value):
                    all_isoc_names.append(cell.value)
            except Exception as e:
                continue
    for cell in perks_workbook['iSOC Perks']['D']:
        try:
            if len(cell.value) > 1 and '>' not in str(cell.value):
                all_perks_names.append(cell.value)
                all_perks_cells.append(cell)
        except Exception as e:
            continue

    for name in all_isoc_names:
        if name not in all_perks_names:
            print(name)


def main():
    token = authentication.authenticate_WiW_API()
    print('Enter Start Date UTC (ex. 04 Feb 2030): ')
    start_date = get_start_date()
    print('Enter End Date: ')
    end_date = get_start_date()
    save_location = f'{str(authentication.get_perks_directory())}cSOC_Perks_{str(start_date.strftime("%b%-d"))}-{end_date.strftime("%b%-d")}'
