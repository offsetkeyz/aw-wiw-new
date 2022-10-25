from distutils import config
import json
import os
import sys
import requests

# Open the configuration json to access API login credentials
path_to_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
if not os.path.isfile(path_to_json):
    sys.exit("'config.json' not found! Please add it and try again.")
else:
    with open(path_to_json) as file:
        config = json.load(file)

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

def get_url_and_headers(type, token):
    url = "https://api.wheniwork.com/2/" + str(type)
    headers = {
    'Host': 'api.wheniwork.com',
    'Authorization': 'Bearer ' + token, 
    }
    return [url, headers]  

def get_schedule_location():
    return config['Excel Schedule Location']

def get_perks_sheet_location():
    return config['Perks Template Location']

def get_perks_directory():
    sb = config['Perks Template Location']
    return str(config['Perks Template Location'].replace("iSOC Perks - Template.xlsx", ""))