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
import pytz
import authentication
import requests
from openpyxl import load_workbook
import datetime
from dateutil.tz import *

schedule_wb = load_workbook('/Users/colin.mcallister/Library/CloudStorage/OneDrive-ArcticWolfNetworksInc/Documents/triage_schedule_from_wiw.xlsx')
perks_workbook = load_workbook('/Users/colin.mcallister/Library/CloudStorage/OneDrive-ArcticWolfNetworksInc/cSOC Perks Sheets/iSOC Perks - Template.xlsx')
save_location = '/Users/colin.mcallister/Library/CloudStorage/OneDrive-ArcticWolfNetworksInc/cSOC Perks Sheets/iSOC Perks Oct9-22_script.xlsx'
perks_ws = perks_workbook['iSOC Perks']

def main():
    token = authentication.authenticate_WiW_API()