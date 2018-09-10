
from __future__ import print_function
import httplib2
import os
import time

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

f1 = open('normalized_peg.dat', 'w')
f2 = open('normalized_eps.dat', 'w')

s =  \
"""QCOM
FB
AMZN
AVGO
AAPL
NFLX
NVDA
GOOGL
IBM
BIDU
GS
BKNG
TSLA
CMG
BABA
INTC
MSFT"""



lines = s.split('\n')
#print(lines)
my_list = lines


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
    spreadsheetId = '1kEETTM4Qt3lHrlWfzJPc6VkdLzYz0kSRi9_zBWPwQHI'
    return spreadsheetId,service


def get_NormalizedEPS():
    
    spreadsheetId, service = get_sheetDetails()
    
    rangeName = 'F98:F98'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_NormalizedEPS:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]
def get_NormalizedPE():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F99:F99'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_NormalizedPE:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            #print('%s' % (row[0]))
            return row[0]

def get_NormalizedPEG():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F95:F95'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_NormalizedPEG:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]

def get_ProfitMargin():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F102:F102'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_ProfitMargin:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]


def get_EquityShare():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F106:F106'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_EquityShare:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]

def get_ExpectedPrice():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F63:F63'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_ExpectedPrice:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]

def get_Price():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F62:F62'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_Price:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]

def get_eps():

    spreadsheetId, service = get_sheetDetails()

    rangeName = 'F94:F94'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('get_EPS:No data found.')
    else:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))
            return row[0]

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

    rangeName = 'A76'
    value_input_option = 'USER_ENTERED'
    #value_input_option = 'QCOM'
    result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId,range = rangeName, valueInputOption=value_input_option, body=body).execute()
    time.sleep(6)

    


def main():
    """Shows basic usage of the Sheets API."""


    #symbool_list = ['QCOM','AAPL','BUD','RIG','MSFT','BABA','SU','YUM','GILD','AMGN','BXP','CSCO','TXN','NVDA']


    #:symbool_list = ['TSLA','INTC','BIIB','IBM','AAPL','QCOM','AVGO','CVS','NVDA','FB','EXPE','GS','AMZN','GOOG']


    #symbol_list = ['QCOM','AAPL']
 

    for symbol in my_list:
        #print ("++++++++++++++++++++++++++++++++")
        print (symbol)
        print()
        update_cell(symbol)
        normPEG = get_NormalizedPEG()
        normPE  = get_NormalizedPE()
        normEPS = get_NormalizedEPS()
        profitMargin  = get_ProfitMargin()
        eps = get_eps()
        ExpPrice = get_ExpectedPrice()
        Price = get_Price()
        Equity = get_EquityShare()
        #print (normEPS)
        #print ("\n++++++++++++++++++++++++++++++++\n")
        #print (profitMargin)
        #print ("\n++++++++++++++++++++++++++++++++\n")
        #print (Equity)
        #print ("\n++++++++++++++++++++++++++++++++\n")
        
        print ("++++++++++++++++++++++++++++++++")
        f1.write("Company\t" + symbol + "\tNORMAL_PEG\t" + str(normPEG) + "\tNORMAL_PE\t"   + str(normPE)  + "\tNORMAL_EPS\t" + str(normEPS) + "\tEQUITY\t" + str(Equity) + "\tPROFIT_MGN\t" + str(profitMargin) +"\n")
        f2.write("Company\t" + symbol + "\tExpPrice\t" + str(ExpPrice)  +  "\tPrice\t" + str(Price) + "\tEPS\t" + str(eps)+ "\tNORMAL_EPS\t" + str (normEPS)+"\n")



if __name__ == '__main__':
    main()

