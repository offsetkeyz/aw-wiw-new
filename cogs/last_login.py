#!/usr/bin/env python3

from dataclasses import field
from datetime import *
import json
import os
import sys
import requests
import csv

token = ''

# Open the configuration json to access API login credentials
path_to_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config-colin.json')
if not os.path.isfile(path_to_json):
    sys.exit("'config.json' not found! Please add it and try again.")
else:
    with open(path_to_json) as file:
        config = json.load(file)

# Class for each User
class User:
    
        def __init__(self, first_name, last_name, email, wiw_employee_id, positions, role, locations, is_hidden, is_active, last_login) -> None:
            self.first_name=first_name
            self.last_name=last_name
            self.email=email
            self.positions=positions
            self.role = role
            self.last_login= last_login
            self.locations = locations
            self.is_hidden = is_hidden
            self.is_active = is_active
            self.full_name = str(first_name) + ' ' + str(last_name)

            if wiw_employee_id == '':
                    self.wiw_employee_id = 0
            else:
                self.wiw_employee_id=int(wiw_employee_id)


def authenticate_WiW_API():

    url = "https://api.wheniwork.com/2/login"
    payload = json.dumps({
    "username": config['username'], 
    "password": config['password'] 
    })
    headers = {
    'W-Key': config['API Key'],
    'Content-Type': 'application/json'
    }
    
    response = requests.request("POST", url, headers=headers, data=payload)
    try:
        token = response.json()['token']
        print("Good to Go!")
        return token
    except:
        print(str(response) + " | Password or API Key Incorrect.")
        return authenticate_WiW_API()

def get_url_and_headers(token):
    url = "https://api.wheniwork.com/2/users"
    headers = {
    'Host': 'api.wheniwork.com',
    'Authorization': 'Bearer ' + token, 
    }
    return [url, headers]  

def get_all_wiw_users(token):
    all_users = {}
    url_headers = get_url_and_headers(token)
    response = requests.request("GET", url_headers[0], headers=url_headers[1])
    all_users_json = response.json()['users']
    for u in all_users_json:
        all_users[u['id']] = (User(u['first_name'], u['last_name'], u['email'], u['id'], u['positions'], u['role'], u['locations'], u['is_hidden'], u['is_active'], u['last_login']))
    return all_users

def main():
    
    results = get_all_wiw_users(authenticate_WiW_API())
    dict_ = {}
    fields = ['email', 'last_login']
    with open('last_login.csv', 'w') as csvfile:
        filewriter = csv.writer(csvfile, delimiter=',')
        for u in results:
            _ = results[u]
            dict_[results[u].email] = results[u].last_login
            try:
                filewriter.writerow([results[u].email, datetime.strptime(results[u].last_login, '%a, %d %b %Y %H:%M:%S %z').strftime('%d %b %Y')])
            except Exception as e:
                continue
    return csvfile


if __name__ == "__main__":
    main()