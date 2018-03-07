
from __future__ import print_function
import httplib2
import os
import time

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_sheetDetails():

    """ Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    """spreadsheetId = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
       rangeName = 'Class Data!A2:E'
    """
    #spreadsheetId = '1TVremAZCmLjXD-q6FbHxdVR5-xMdt73O4qXMEXodKF8'
    spreadsheetId = '1SZma3WDG8UrNrvDYT7MErSXlcpLa0VPbWwQlMBkeN6c'
    return spreadsheetId,service


def get_data():
    
    spreadsheetId, service = get_sheetDetails()
    
    rangeName = 'E60:60'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[1]))

def update_cell(symbol):
    
    spreadsheetId, service = get_sheetDetails()
    cell = symbol
    
    values = [
        [
            # Cell values ...
            cell
        ],
        # Additional rows ...
    ]
    #values = [ 'QCOM' ]

    body = {
        'values': values
    }

    rangeName = 'A68'
    value_input_option = 'USER_ENTERED'
    #value_input_option = 'QCOM'
    result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId,range = rangeName, valueInputOption=value_input_option, body=body).execute()
    time.sleep(6)

    


def main():
    """Shows basic usage of the Sheets API."""


    #symbool_list = ['QCOM','AAPL','BUD','RIG','MSFT','BABA','SU','YUM','GILD','AMGN','BXP','CSCO','TXN','NVDA']


    #symbool_list = ['TSLA','INTC','BIIB','IBM','INFY','SWKS','TGT','M','BIDU','FB','JPM','PCLN','KSS','C']


    symbool_list = ['QCOM','AAPL','AVGO','WMT']
 

    for symbol in symbool_list:
        print ("++++++++++++++++++++++++++++++++")
        print (symbol)
        update_cell(symbol)
        get_data()
        print ("++++++++++++++++++++++++++++++++")

if __name__ == '__main__':
    main()

