#!/usr/bin/env python3
""""
Author: Colin McAllister (https://github.com/offsetkeyz)

This script will query all users from When I Work instance and return a CSV of their last log-in.
"""
import sys


from dataclasses import field
from datetime import *
import authentication
import requests
import csv

token = authentication.authenticate_WiW_API()

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

def get_all_wiw_users(token):
    all_users = {}
    url_headers = authentication.get_url_and_headers(token)
    response = requests.request("GET", url_headers[0], headers=url_headers[1])
    all_users_json = response.json()['users']
    for u in all_users_json:
        all_users[u['id']] = (User(u['first_name'], u['last_name'], u['email'], u['id'], u['positions'], u['role'], u['locations'], u['is_hidden'], u['is_active'], u['last_login']))
    return all_users

def main():
    
    results = get_all_wiw_users(token)
    dict_ = {}
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