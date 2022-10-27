#TODO add INACTIVE to inactive users.
#TODO add file for classes.

import datetime
from openpyxl import load_workbook
import pytz
import requests
from cogs import classes, authentication
from openpyxl.styles import Font, Color, PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
import csv
from dateutil.tz import *
import re
from datetime import *
import json

token = ''

excel_schedule_workbook = load_workbook(authentication.get_schedule_location())
date_columns = {}
all_names = {}
date_row = 2

def get_url_and_headers(type):
    url = "https://api.wheniwork.com/2/" + str(type)
    headers = {
    'Host': 'api.wheniwork.com',
    'Authorization': 'Bearer ' + token, 
    }
    return [url, headers] 

def build_date_row():
    for sheet in excel_schedule_workbook.sheetnames:
        current_ws = excel_schedule_workbook[sheet]
        start_date = datetime(2022,2,28)
        for col in current_ws.iter_cols(min_row=date_row, max_row=date_row, min_col=3, max_col=500):
            for cell in col:
                cell.value = datetime.strftime(start_date, '%d %b %Y')
                date_columns[datetime.strftime(start_date, '%d %b %Y')] = cell.column #stores columns of dates in global dict. Columns same for all sheets
                cell.fill = PatternFill("solid", fgColor="FFEFC2")
                if datetime.strftime(datetime.strptime(cell.value, '%d %b %Y'), '%a').startswith('S'): #weekend
                    cell.fill = PatternFill("solid", fgColor="AFD5FF")
            start_date = start_date + timedelta(days=1)
        # excel_schedule_workbook.save(str(workbook_location))

def clear_future_columns(start_date:datetime):
    get_date_row()
    for i in excel_schedule_workbook.sheetnames:
        ws = excel_schedule_workbook[i]
        ws.delete_cols(date_columns[datetime.strftime(start_date, '%d %b %Y')], ws.max_column)
    build_date_row()
    get_date_row()

def get_date_row():
    current_ws = excel_schedule_workbook['TSE1']
    for cell in current_ws[date_row]:
        date_columns[str(cell.value)] = cell.column #stores columns of dates in global dict. Columns same for all sheets

# returns dict with Worksheet as key and an array of names as value
def get_all_names():
    for sheet in excel_schedule_workbook.worksheets:
        sheet_names = {} # names : column number
        for cell in sheet['A']:
            sheet_names[cell.value] = cell.row
        all_names[sheet.title] = sheet_names

def populate_excel_sheet(shifts_in):
    for user_id in shifts_in:
        user = get_user_from_id(user_id)
        all_to_requests = get_time

# ------------------------- WHEN I WORK METHODS --------------------------- #

# gets all shifts as shift classes and stores in dictionary by user
def get_all_shifts():    
    url_headers = get_url_and_headers('shifts?start=' + str(datetime(2022, 3, 1)) + "&end=" + str(datetime.now()+ timedelta(days=360)) + "&unpublished=true")
    response = requests.request("GET", url_headers[0], headers=url_headers[1])
    all_shifts = response.json()['shifts']
    employee_shifts = {}
    for i in all_shifts:
        start_time = datetime.strptime(i['start_time'], '%a, %d %b %Y %H:%M:%S %z').astimezone(pytz.timezone('UTC'))
        end_time = datetime.strptime(i['end_time'], '%a, %d %b %Y %H:%M:%S %z').astimezone(pytz.timezone('UTC'))
        new_shift = classes.shift(int(i['id']), int(i['account_id']), int(i['user_id']), int(i['location_id']), int(i['position_id']),
                                int(i['site_id']), start_time, end_time, bool(i['published']), bool(i['acknowledged']), i['notes'], i['color'], bool(i['is_open']))
        if int(i['user_id']) in employee_shifts:
            current_users_shifts = employee_shifts.get(int(i['user_id']))
            current_users_shifts.append(new_shift)
            employee_shifts[int(i['user_id'])] = current_users_shifts
        else: 
            employee_shifts[int(i['user_id'])] = [new_shift]
    return employee_shifts

def get_time_off_requests_for_user(user_id):
    url_headers = get_url_and_headers('requests?start=' + str(datetime(2022, 3, 1)) + "&end=" + str(datetime.now()+timedelta(days=100)) + '&user_id=' + str(user_id) + '&limit=200', token)
    try: #permissions error
        response = requests.request("GET", url_headers[0], headers=url_headers[1])
        all_requests = response.json()['requests']
    except Exception as e:
        all_requests = {}
    return store_time_off(all_requests)

def get_all_wiw_users():
    all_users = {}
    url_headers = get_url_and_headers('users')
    response = requests.request("GET", url_headers[0], headers=url_headers[1])
    all_users_json = response.json()['users']
    for u in all_users_json:
        all_users[u['id']] = (classes.user(u['first_name'], u['last_name'], u['email'], u['id'], u['positions'], u['role'], u['locations'], u['is_hidden'], u['is_active']))
    return all_users

#takes in user_id and returns user object
def get_user_from_id(user_id):
    url_headers = get_url_and_headers('users/'+str(user_id))
    i = 1
    while i < 10:
        try:
            j = requests.request("GET", url_headers[0], headers=url_headers[1]).json()['user']
            break
        except:
            i +=1 
    return classes.user(j['first_name'], j['last_name'], j['email'], j['id'],j['positions'],j['role'],j['locations'],j['is_hidden'],j['is_active'])


def main():
    token = authentication.authenticate_WiW_API()
    build_date_row()
    clear_future_columns(datetime.now()-timedelta(days=14))
    get_all_names()
    populate_excel_sheet(get_all_shifts())
