#!/usr/bin/env python3
""""
Author: Colin McAllister (https://github.com/offsetkeyz)

This script will duplicate a when I work user's schedule froma  certain date.
"""

from datetime import timedelta
import datetime
import json
import random
from time import strptime
import pytz
from cogs import authentication, classes
import requests

def get_email():
    user_name = input().strip()
    if not user_name.endswith('@arcticwolf.com'):
        print("Incorrect Format with username: " + user_name)
        print('Please try again: ')
        user_name=get_email()
    return user_name

def get_start_date():
    input_date = input().strip()
    try:
        dt_start = datetime.datetime.strptime(str(input_date), '%d %b %Y')
        dt_start = dt_start.replace(hour=13)
    except ValueError as e:
        print ("Incorrect format with get_start_date: " + input_date)
        print('Please try again (ex. 07 Jul 2033): ')
        input_date = get_start_date()
    return dt_start

def get_user_id_from_email(token, user_email):
    url_headers = authentication.get_url_and_headers('users?search='+user_email, token)
    success = False
    i = 1
    while success == False & i < 10:
        try:
            response = requests.request("GET", url_headers[0], headers=url_headers[1])
            success = True
        except:
            success = False
            i +=1 
    try:
        user_id = response.json()['users'][0]['id']
    except:
        print(f'Error with user email: {user_email}')
        user_id = 0
    return user_id

def get_team_id(team_number):
    try:
        n = int(team_number)
    except:
        print("Team number not INT")
    teams = {
        1: 4781862,2: 4781869,3: 4781872,4: 4781873,5: 4781874,6: 4781875,7: 4781876,8: 4781877,9: 4781878,10: 4781879
     }
    try:
         return teams[n]
    except: 
        return team_number

def get_color_code(color):
    shift_color = color
    if color == "red":
        shift_color = "FF0000"
    elif color == "blue" or color == "dark blue":
        shift_color = "0070c0"
    elif color == "purple":
        shift_color = "8d3ab9"
    elif color == "orange":
        shift_color = "FFC000"
    elif color == "teal":
        shift_color = "72F2DA"
    elif color == "green":
        shift_color = "42a611"
    elif color == "gray":
        shift_color = "a6a6a6"
    elif color == "yellow":
        shift_color = "ffff00"
    elif color == "light blue":
        shift_color = "00b0f0"
    elif color == "pink":
        shift_color = "ff00dd"
    elif color == "black":
        shift_color = "000000"
    elif color == "MST": 
        shift_color = 'A1FECF'
    elif color == "CST":
        shift_color = "32B272"
    elif color == "EST":
        shift_color = '007038'
    return shift_color

def get_all_future_shifts_json(token, start_date):    
# url_headers = get_url_and_headers('shifts')
    url_headers = authentication.get_url_and_headers('shifts?start=' + str(start_date) + "&end=" + str(datetime.datetime.now()+ timedelta(days=360)) + "&unpublished=true", token)
    response = requests.request("GET", url_headers[0], headers=url_headers[1])
    all_shifts = response.json()['shifts']
    return all_shifts

def get_all_shifts_json(token):    
# url_headers = get_url_and_headers('shifts')
    url_headers = authentication.get_url_and_headers('shifts?start=' + str(datetime(2022, 3, 1)) + "&end=" + str(datetime.datetime.now()+ timedelta(days=360)) + "&unpublished=true", token)
    response = requests.request("GET", url_headers[0], headers=url_headers[1])
    all_shifts = response.json()['shifts']
    return all_shifts

def store_shifts_by_user_id(all_shifts_in):
        # key: user_id | value: array of shifts
    employee_shifts = {}
    for i in all_shifts_in:
        start_time = datetime.datetime.strptime(i['start_time'], '%a, %d %b %Y %H:%M:%S %z').astimezone(pytz.timezone('UTC'))
        end_time = datetime.datetime.strptime(i['end_time'], '%a, %d %b %Y %H:%M:%S %z').astimezone(pytz.timezone('UTC'))
        new_shift = classes.shift(int(i['id']), int(i['account_id']), int(i['user_id']), int(i['location_id']), int(i['position_id']),
                                int(i['site_id']), start_time, end_time, bool(i['published']), bool(i['acknowledged']), i['notes'], i['color'], bool(i['is_open']))
        if int(i['user_id']) in employee_shifts:
            current_users_shifts = employee_shifts.get(int(i['user_id']))
            current_users_shifts.append(new_shift)
            employee_shifts[int(i['user_id'])] = current_users_shifts
        else: 
            employee_shifts[int(i['user_id'])] = [new_shift]
    return employee_shifts

def delete_all_shifts_for_user(token, start_date, user_id, all_shifts=0):
    def delete_shift(shift_id, token):
        url_headers = authentication.get_url_and_headers('shifts/' + str(shift_id), token)
        response = requests.request("DELETE", url_headers[0], headers=url_headers[1])

    start_date = start_date - timedelta(hours=15)
    if all_shifts == 0: #option to pass in dictionary of shift\=]-
        all_shifts_json = get_all_shifts_json(token)
        all_shifts = store_shifts_by_user_id(all_shifts_json)
    try: 
        for shift in all_shifts[user_id]:
            try:
                if shift.start_time >= start_date.astimezone(pytz.timezone('UTC')):
                    delete_shift(shift.shift_id, token)
            except Exception as f:
                if shift.start_time >= pytz.timezone('UTC').localize(start_date):
                    delete_shift(shift.shift_id, token)
    except Exception as e:
        print("I'm so miffed, this user has no shifts")

def copy_users_schedule(user_id_to_copy, new_user_email, start_date, token):
    def create_duplicate_shift(token, user_email, start_time, length, color, notes, schedule_id, team_number, position):
        user_id = get_user_id_from_email(token, user_email)    
        start_hour = int(start_time.strftime('%H'))
        start_time = start_time.replace(hour=(start_hour))
        end_time = start_time + timedelta(hours=length) 
        shift_color = get_color_code(color)
        url_headers = authentication.get_url_and_headers('shifts', token)
        payload = json.dumps({
            "user_id": user_id, 
            "location_id": schedule_id,
            "start_time": start_time.strftime("%a, %d %b %Y %H:%M:%S %z"),
            "end_time" : end_time.strftime("%a, %d %b %Y %H:%M:%S %z"),
            "color": shift_color,
            "notes" : notes,
            "site_id" : get_team_id(team_number),
            "position_id": position
        }) 
        success = False
        i = 1
        while success == False & i < 10:
            try:   
                request = requests.request("POST", url_headers[0], headers=url_headers[1], data=payload)
                success = True
            except:
                success == False
                i += 1

    all_shifts_json = get_all_future_shifts_json(token, start_date)
    all_shifts = store_shifts_by_user_id(all_shifts_json)
    print(f"Deleting {new_user_email}'s old shifts")
    delete_all_shifts_for_user(token, start_date, get_user_id_from_email(token,new_user_email), all_shifts)
    print(f"Copying shifts...")
    i = 0
    for shift in all_shifts[user_id_to_copy]:
        if shift.start_time >= start_date.astimezone(pytz.timezone('UTC')):
            create_duplicate_shift(token, new_user_email, shift.start_time, shift.length, shift.color, shift.notes, shift.location_id, shift.site_id, shift.position_id)
            i = i+1
    print(f'Copied {i} shifts to {new_user_email}.')


def main():
    token = authentication.authenticate_WiW_API()
    print('Please enter the email of the user you are copying: ')
    user_to_copy = get_email()
    print('Enter the email of the user getting a new schedule: ')
    user_to_paste = get_email()
    print('Enter the date to start copying (DD mmm YYYY) example: 06 Sep 2025')
    start_date = get_start_date()
    print(f"You are copying {user_to_copy}'s schedule and pasting onto {user_to_paste}'s schedule starting on {start_date}. Is this correct? (Y/N): ")
    r = input()
    if r.strip().lower() not in ['y', 'yes']:
        print("ok bye")
        return
    print(random.choice(["Ok loozer, lets go.", "Why do you make me work so hard.", "Please contact IT for support. JK I'M GOING OK?!", "LFG", "rip in pieces", "peepee poopoo they're lucky to have you.", "Please put on a movie while you wait. You deserve it queen."]))
    copy_users_schedule(get_user_id_from_email(token, user_to_copy), user_to_paste, start_date, token)

if __name__ == "__main__":
    main()