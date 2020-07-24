
# It get a list of Messages from the user's mailbox then
# make download of preciocierresTR.xls file to selected date

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
from Config import localDir

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists(localDir + '/token.pickle'):
    with open(localDir + '/token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            localDir + '/credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(localDir + '/token.pickle', 'wb') as token:
        pickle.dump(creds, token)


def GetAttachments(service, user_id, msg_id, pathfile):
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    # pprint(message)
    for part in message['payload']['parts']:
        if part['filename']:
            if 'data' in part['body']:
                data = part['body']['data']
                
            else:
                att_id = part['body']['attachmentId']
                att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
                data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            
            filename = part['filename']
            if filename == 'preciocierresTR.xls':
                # print(part['filename'])
                with open(pathfile +'/' + filename, 'wb') as f:
                    f.write(file_data)
    return pathfile +'/' + filename

def  ListMessagesMatchingQuery(service, user_id, query=''):
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

def getAttcEmailGmail(date, pathRoot):
    dtAfter = datetime.strftime(datetime.strptime(date, '%Y-%m-%d'), '%m/%d/%Y')
    dtBefore = datetime.strftime(datetime.strptime(date, '%Y-%m-%d') + timedelta(days=1), '%m/%d/%Y')
    service = build('gmail', 'v1', credentials=creds)

    if dtAfter == datetime.strftime(datetime.now(), '%m/%d/%Y'):
        a = ListMessagesMatchingQuery(service,'<EMAIL GMAIL ACCOUNT>','label:Panama after:'+dtAfter)
        # print(dtAfter)
    else:
        a = ListMessagesMatchingQuery(service,'<EMAIL GMAIL ACCOUNT>','label:Panama after:'+dtAfter+' before:'+dtBefore) 
        # print(dtAfter, dtBefore)
    # pprint(a)
    
    if len(a) > 1:
        pathFiles = []
        for a1 in range(len(a)):
            pathFile = GetAttachments(service, 'me', a[a1]['id'], pathRoot)
            pathFiles.append(pathFile)
        return pathFiles

    elif len(a) == 1:
        for a1 in range(len(a)):
            pathFile = GetAttachments(service, 'me', a[a1]['id'], pathRoot)
        return pathFile
    else:
        return 0


#x = getAttcEmailGmail('2020-05-14', 'C:/Users/u6081428/Projects/Panama')