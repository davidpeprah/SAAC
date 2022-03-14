#
#
#  Author: David Peprah
#
#

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from configparser import ConfigParser
import subprocess, sys

# Custom functions from lib folder
sys.path.insert(0, 'lib')
from checkUserGS import checkUser
from sendEmail import sendMessage, CreateMessage


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/admin.directory.user.readonly', 'https://www.googleapis.com/auth/gmail.send']

# Load Config file
config = ConfigParser()
config.read('config/config.ini')


if 'Document' not in config.sections() or 'admin' not in config.sections():
    exit()

# The ID and range of a sample spreadsheet.
if (config.get('Document', 'SpreadSheetID')):
    SPREADSHEET_ID = config.get('Document', 'SpreadSheetID')
    SHEETS = 'Sheet1' if not (config.get('Document', 'Sheets')) else config.get('Document', 'Sheets')
    Range = '!A2:R' if not (config.get('Document', 'SheetRange')) else config.get('Document', 'SheetRange')
    EntryTypeColAdd = config.get('Document', 'EntryTypeColAdd')
    NewMailColAdd = config.get('Document', 'NewMailColAdd')
    StatusColAdd = config.get('Document', 'StatusColAdd')
    CommentColAdd = config.get('Document', 'CommentColAdd')
else:
    exit()

# Get authorized users and domains
if (config.get('admin', 'AuthorizeUsers')):
    authUsers = list(config.get('admin', 'AuthorizeUsers').split(','))
    domain = config.get('admin', 'Domain')
    admin = config.get('admin', 'sysadmin')
    openticket = config.get('admin', 'openticket')
    srvAccEmail = config.get('admin', 'serviceAccEmail')
else:
    exit()

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
    if os.path.exists(r'credentials\token.pickle'):
        with open(r'credentials\token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r'credentials\credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(r'credentials\token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)


    # Call the Sheets API
    sheet = service.spreadsheets()


def readSheet(sh):
    global sheet
    
    RANGE_NAME = sh + Range

    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    #print(result)
    if not values:
        print('No data found.')
    else:
        row_count = 2
        for row in values:
            try:
                if row[8] == "NEW":
                    
                    currentEmpEmail = '"' + row[1] + '"'
                    lastName = '"' + row[2].title() + '"'
                    firstName = '"' + row[3].title() + '"'
                    #lastFourSSN = '"' + row[2] + '"'
                    personalEmail = '"' + row[4] + '"'
                    position = '"' + row[5].title() + '"'
                    building = '"' + row[6] + '"'
                    department = '"' + row[7] + '"'

                    #if (department in districtDepart) and (building == 'District'):

                    if currentEmpEmail.split("@")[0] in authUsers:
                        # Send Data to Powershell
                        createAcc = subprocess.Popen(["Powershell.exe", r'lib\setAcc.ps1 ' +\
                        currentEmpEmail + ' ' + lastName + ' ' + firstName + ' ' + personalEmail + ' ' + str(position) + ' ' +\
                        str(building) + ' ' + department ], stdout=subprocess.PIPE)

                        # Read Information from Powershell
                        message = str(createAcc.communicate()[0][:-2], 'utf-8')
                        status, email, update = message.split('\r\n')


                        if (status == "1"):
                            UpdateStatus(sh,row_count,status_msg(status),update)
                            UpdateMail(sh,row_count,email)
                            UpdateEntryType(sh,row_count,message)
                        
                        elif (status == "2"):
                            UpdateStatus(sh,row_count,status_msg(status),update)
                            UpdateEntryType(sh,row_count,message)

                            # Send an email to sysaid and copy ERC Team when there is an error
                            fname = firstName.strip('"')
                            lname = lastName.strip('"')
                            
                            ticketMsg = f"Hi Tech Team\n\nAn error occured while creating an account for the user below: \nFullname: {fname} {lname}\nPlease check the event log which is located in the log folder of this application for details\n\nThank you"

                            ticketSubj = 'Account Creation Error - Hamilton'
                            sendMessage('me', CreateMessage(srvAccEmail, openticket, ticketSubj,ticketMsg,admin))
                                                     
                            
                                                    
                    else:
                        UpdateStatus(sh,row_count,status_msg("3"),"Access Denied")
                        UpdateEntryType(sh,row_count,"OLD")
                        
                        unauthorizedUser = row[1]
                        deniedAccessMsg =  f"Hi\n\n  Please you are not authorized to create district email account. \n Contact the Human Resource department or send an email to {authUsers[0] + '@' + domain} \n\n Thank you"
                        deniedSubj = 'UNAUTHORIZED USER - Hamilton'
                        # Send an email to the unauthorized user and copy the admin if further investigation needed
                        sendMessage('me', CreateMessage(srvAccEmail, unauthorizedUser, deniedSubj,deniedAccessMsg, admin))
                                                                             

                elif row[10] == "Pending":
                    email = str(row[9])
                    personal_email = str(row[4])
                    firstName = str(row[3])
                    lastName = str(row[2])
                    CurEmpEmail = str(row[1])
                    
                    try:
                        check = checkUser(email)
                        
                        if check == email:
                            UpdateStatus(sh,row_count,status_msg("0"),"Account Successfully confirm in G-Suite")
                            passw = password()
                            passReset = subprocess.Popen(["Powershell.exe", r'lib\resetPass.ps1 ' + email + ' ' + '"' + passw + '"' ], stdout=subprocess.PIPE)
                            msg = str(passReset.communicate()[0][:-2], 'utf-8')
                            status, user = msg.split('\r\n')

                            
                            #Send account information to the new staff personal email and notify the employee who made the entry
                            if (personal_email):
                                
                                newHireMsg = f"Hi {firstName}\n\nPlease your account information is provided below:\n\tusername: {user}\n\temail: {email}\n\tpassword: {passw}\n\nKindly send an email to {admin} if you need any assistance\n\nThank you"
                                newHireSubj = f'Account Information for  {firstName} {lastName}'
                                sendMessage('me', CreateMessage(sender=srvAccEmail, to=personal_email, subject=newHireSubj,message_text=newHireMsg))
                                

                                empMsg  = f"Hi \n\nThis is to notify you that the account for {firstName} {lastName} was created successfully.\nAn email containing the account information has been sent to the new hire using the personal email you provided.\nThank you"
                                empSubj = f'Account for {firstName} {lastName}  Completed Successfully'
                                sendMessage('me', CreateMessage(sender=srvAccEmail, to=CurEmpEmail, subject=empSubj, message_text=empMsg))
                            
                    except HttpError as err:
                        if err.resp.status in [404,]:
                            UpdateStatus(sh,row_count,status_msg("1"),"Account has not been created in Google Console yet")
                        else:
                            UpdateStatus(sh,row_count,status_msg("2"),"Error occured when trying to verify account in Google console")

                else:
                    print("Nothing to work on!!!")


                #elif row[16] == "Active":
                    


                #elif row[16] == "Activiting"
                    

            except IndexError:
                row_count += 1
                continue
            
            row_count += 1

def UpdateStatus(sh,address,status,update):
    global sheet

    RANGE_NAME = sh + "!" + StatusColAdd + str(address) + ":" + CommentColAdd + str(address)

    body_value = {"majorDimension": "ROWS", "values": [[str(status),str(update)]]}

    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()

    #values = result.get('values', [])
    #print(result)

def UpdateMail(sh,address,message):
    global sheet

    RANGE_NAME = sh + "!" + NewMailColAdd + str(address)

    body_value = {"majorDimension": "ROWS", "values": [[str(message)]]}


    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption='USER_ENTERED', body=body_value).execute()

    #values = result.get('values', [])
    #print(result)

def UpdateEntryType(sh,row_add,message):
    global sheet
    
    RANGE_NAME = sh + "!" + EntryTypeColAdd + str(row_add)

    body_value = {"majorDimension": "ROWS", "values": [[str("OLD")]]}

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
    elif status == "3":
        return "Denied"

# Returns year if last four digit of social security is not found
def checkSSN(ssn):
    if len(ssn) == 4:
        return ssn
    else:
        import datetime
        year = datetime.datetime.now()
        year.strftime("%Y")
        return year


def password():
    import random
    import array

    # Maximum length of password 
    MAX_LEN = 12

    # An empty List to hold all the characters
    COMBINED_LIST = []

    # At least one of these characters should be in the password
    # The characters are picked randomly. This uses ASCII 
    DIGITS = chr(random.randint(48, 57))
    UPPERCASE = chr(random.randint(65, 90))
    LOWERCASE = chr(random.randint(97, 122))
    SYMBOLS = ['&', '@', '!', '*', '%', '#']
    RANDSYM = SYMBOLS[random.randint(0,5)]
    

    # Combined all the defined characters into a list
    for x in range(48, 57): COMBINED_LIST += chr(x)
    for x in range(97, 122): COMBINED_LIST += chr(x)
    for x in range(65, 90): COMBINED_LIST += chr(x)
    for x in SYMBOLS: COMBINED_LIST += x

    # Combine the initial randomly selected characters from above
    temp_pass = DIGITS + UPPERCASE + LOWERCASE + RANDSYM

    # Add additional 8 characters to make up for the 12 Maximum length
    for x in range(MAX_LEN -4):
        temp_pass += random.choice(COMBINED_LIST)

        # convert temporary password into array and shuffle to
        # prevent it from having consistent pattern
        # where the beginning of the password is predictable
        temp_pass_list = array.array('u', temp_pass)
        random.shuffle(temp_pass_list)

    # traverse the temporary password array and append the chars
    # to form the password
    password = ""
    for x in temp_pass_list:
        password += x

    return(password)

    

if __name__ == '__main__':
    # call the main function 
    main()

    # Loop through the sheet
    SHEETS = list(SHEETS.split(','))
    for sh in SHEETS:
        readSheet(sh)
