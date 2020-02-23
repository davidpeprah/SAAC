from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/admin.directory.user.readonly']

def checkUser(email):
    """Shows basic usage of the Admin SDK Directory API.
    Prints the emails and names of the first 10 users in the domain.
    """
    
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token1.pickle'):
        with open('token1.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials1.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token1.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('admin', 'directory_v1', credentials=creds)

    # Call the Admin SDK Directory API
    results = service.users().get(userKey=email, viewType='domain_public').execute()
    user = results
    

    if not user:
        return ('No users in the domain.')
    else:
        return (u'{0}'.format(user['primaryEmail']))

