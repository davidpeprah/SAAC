from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

import subprocess, sys
from checkUserGS import checkUser
from sendEmail import sendMessage, CreateMessage

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1gKZCaLZq79LzwxB_wpHpU6Tu4R-eTr5vGyLFG_ZQ0R8'
sheet = ""


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    global sheet
    
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    

    # Call the Sheets API
    sheet = service.spreadsheets()
    

def readSheet():
    global sheet
    
    RANGE_NAME = 'Sheet1!A2:I'

    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    #print(result)
    if not values:
        print('No data found.')
    else:
        row_add = 2
        for row in values:
            if row[8] == "New":
                lastName = '"' + row[0] + '"'
                firstName = '"' + row[1] + '"'
                lastFourSSN = '"' + row[2] + '"'
                position = '"' + row[5] + '"'
                building = '"' + row[6] + '"'
                staffType = '"' + row[7] + '"'
                createAcc = subprocess.Popen(["Powershell.exe", r'c:\users\peprah_d\desktop\AutoAccProject\setAcc.ps1 ' + lastName + ' ' + firstName + ' ' + lastFourSSN + ' ' + str(position) + ' ' + str(building) + ' ' + staffType ], stdout=subprocess.PIPE)
                message = createAcc.communicate()[0][:-2]
                status, email, update = message.split('\r\n')
               

                if (status == "1"):
                    UpdateStatus(row_add,status_msg(status),update)
                    UpdateMail(row_add,email)
                elif (status == "2"):
                    UpdateStatus(row_add,status_msg(status),update)

            elif row[8] == "Pending":
                email = str(row[4])
                personal_email = str(row[3])
                
                try:
                    check = checkUser(email)
                    if check == email:
                        UpdateStatus(row_add,status_msg("0"),"Account Successfully confirm in G-Suite")
                        lastFourSSN = str(row[2])
                        passReset = subprocess.Popen(["Powershell.exe", r'c:\users\peprah_d\desktop\AutoAccProject\resetPass.ps1 ' + email + ' ' + lastFourSSN ], stdout=subprocess.PIPE)
                        msg = passReset.communicate()[0][:-2]

                        #Send account information to staff and copy diane hill
                        if (personal_email):
                            sendMessage('me', CreateMessage('peprah_d@milfordschools.org', personal_email, 'dpeprah@vartek.com', 'Account Information for ' + row[1] + ' ' + row[0], \
                                                            'Hi ' + row[1] + '\n\n' + \
                                                            'Please your account information has been provided below:\n' +\
                                                            ' username: '+ email.split('@')[0] + '\n' +\
                                                            ' email:    '+ email + '\n' +\
                                                            ' password: milford'+ checkSSN(lastFourSSN) +'\n\n\n'   +\
                                                            'Kindly reply to this mail if you need any assistance\n\n' +\
                                                            'Thank you'))
                except HttpError as err:
                    if err.resp.status in [404,]:
                        UpdateStatus(row_add,status_msg("1"),"Account have not been created in Google Console yet")
                    else:
                        UpdateStatus(row_add,status_msg("2"),"Error occured when trying to verify account in Google console")
                    
                
            row_add += 1

def UpdateStatus(address,status,update):
    global sheet

    RANGE_NAME = "Sheet1!I" + str(address) + ":J" + str(address)

    body_value = {"majorDimension": "ROWS", "values": [[str(status),str(update)]]}
    
    
    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()
    
    #values = result.get('values', [])
    #print(result)

def UpdateMail(address,message):
    global sheet

    RANGE_NAME = "Sheet1!E" + str(address)

    body_value = {"majorDimension": "ROWS", "values": [[str(message)]]}
    
    
    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()
    
    #values = result.get('values', [])
    #print(result)

def status_msg(status):
    if status == "0":
        return "Done"
    elif status == "1":
        return "Pending"
    elif status == "2":
        return "Error"

# Returns year if last four digit of social security is not found
def checkSSN(ssn):
    if len(ssn) == 4:
        return ssn
    else:
        import datetime
        year = datetime.datetime.now()
        year.strftime("%Y")
        return year
    
    
    
if __name__ == '__main__':
    main()
    readSheet()
