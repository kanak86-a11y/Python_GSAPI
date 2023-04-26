from googleapiclient.discovery import build
from google.oauth2 import service_account
from gspread import authorize
import logging
import sys,os
from urllib import request

class Googlesheetsapi2:
    def __init__(self, service_account_file, scopes, spreadsheet_id, range_name):
        self.service_account_file = service_account_file
        self.scopes = scopes
        self.spreadsheet_id = spreadsheet_id
        self.range_name = range_name
        self._creds = None
        self._gclient = None
        self._sheet = None
        self._connect()
    
    def _connect(self):
        creds = service_account.Credentials.from_service_account_file(self.service_account_file, scopes = self.scopes)
        self._creds = creds
        self._gclient = authorize(creds)
        self._sheet = build('sheets', 'v4', credentials=creds).spreadsheets()
        
    def is_connected(self):
        try:
            request.urlopen('https://docs.google.com/spreadsheets/d/1hzv14EF-TppuBWQAVYQ9s2kcHi12yPJH5oVcPqGfM0c/edit#gid=0',timeout=1)
            return True
        except request.urlopen as err:
            return False
        
    def write_data2121(self, data):
        rs = {'values': data}
        updated = self._sheet.values().append(spreadsheetId=self.spreadsheet_id, range=self.range_name, valueInputOption="USER_ENTERED", body=rs).execute() 
        if updated:
            return 'data is updated'
        else:
            return 'no data is updated'

        
    def read_data(self):
        data = self._sheet.values().get(spreadsheetId=self.spreadsheet_id, range=self.range_name).execute()
        return data.get('values', [])
    
    def call_sheet_api(data):
        service_account_file = 'leafy.json'
        scopes = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
        spreadsheet_id = '1hzv14EF-TppuBWQAVYQ9s2kcHi12yPJH5oVcPqGfM0c'
        range_name = 'sheet1'
        sheets_api = Googlesheetsapi2(service_account_file, scopes, spreadsheet_id, range_name)
        if sheets_api.is_connected():
            sheets_api.write_data2121(data)
            data = sheets_api.read_data()
            return '/n/n connected'
        else:
            return 'not connected'

