
# It get a list of Messages from the user's mailbox then
# extract information from the body of the email the maximum 
# number of emails existing on the selected day.

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from pprint import pprint
import base64
from datetime import datetime, timedelta

localDir= os.getcwd().replace('\\','/')

def ListMessagesMatchingQuery(service, user_id, query=''):
    """List all Messages of the user's mailbox matching the query.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

    Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
    """
    response = service.users().messages().list(userId=user_id,
                                                q=query).execute()
    messages = []
    if 'messages' in response:
        messages.extend(response['messages'])

    while 'nextPageToken' in response:
        page_token = response['nextPageToken']
        response = service.users().messages().list(userId=user_id, q=query,
                                            pageToken=page_token).execute()
        messages.extend(response['messages'])

    return messages


def GetMessage(service, user_id, msg_id):
    """Get a Message with given ID.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

    Returns:
    A Message.
    """
    
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()


    return message
    

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists(localDir+'/token.pickle'):
    with open(localDir+'/token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            localDir+'/credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(localDir+'/token.pickle', 'wb') as token:
        pickle.dump(creds, token)


def getEmail(date):
    dtAfter = datetime.strftime(datetime.strptime(date, '%d-%b-%y'), '%m/%d/%Y')
    dtBefore = datetime.strftime(datetime.strptime(date, '%d-%b-%y') + timedelta(1), '%m/%d/%Y')

    service = build('gmail', 'v1', credentials=creds)

    if dtAfter == datetime.strftime(datetime.now(), '%m/%d/%Y'):
        a = ListMessagesMatchingQuery(service,'<EMAIL GMAIL ACCOUNT>','label:Caracas after:'+dtAfter)
    else:
        a = ListMessagesMatchingQuery(service,'<EMAIL GMAIL ACCOUNT>','label:Caracas after:'+dtAfter+' before:'+dtBefore) 

    result = []


    for a1 in range(len(a)):
        b = GetMessage(service, 'me', a[a1]['id'])

        c = b['payload']
        
        listHeaders = c['headers']
        date = ''
        for h in listHeaders:
            dt = False
            for key, value in h.items():
                if dt == False:
                    if value == 'Date':
                        dt = True
                else:
                    date = datetime.strftime(datetime.strptime(value, '%a, %d %b %Y %X +0000'), '%d-%b-%y')

        d = c['parts'][0]

        try:
            e = d['parts'][0]
            z = e['body']['data']
        except:
            z = d['body']['data']

        x = base64.b64decode(z, '-_')
        x1 = x.decode('UTF-8')
        x2 = x1.split('\r\n')
        
        result.append({'id': a[a1]['id'], 'msg':x2, 'date': date})
    return result
        