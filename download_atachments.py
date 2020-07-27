# For gmail API
from __future__ import print_function
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
import base64

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def list_messages_matching_query(service, user_id, query=''):
    """
    List all Messages of the user's mailbox matching the query.

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
    try:
        response = service.users().messages().\
            list(userId=user_id, maxResults=5, q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().\
                list(userId=user_id, q=query, pageToken=page_token).execute()
            messages.extend(response['messages'])

        return messages

    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def get_attachment_by_msg_id(service, user_id, msg_id, store_dir):
    """
    Get and store attachment from Message with given id, taken from gmail API's examples
    :param service: Authorized Gmail API service instance.
    :param user_id: User's email address. The special value "me" can be used to indicate
            the authenticated user.
    :param msg_id: ID of Message containing attachment.
    :param store_dir: The directory used to store attachments.
    :return:
    """
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        for part in message['payload']['parts']:
            if part['filename']:
                attachment = service.users().messages().attachments().\
                    get(userId='me', messageId=message['id'],
                        id=part['body']["attachmentId"]).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                path = ''.join([store_dir, part['filename']])

                f = open(path, 'wb')
                f.write(file_data)
                f.close()

    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def get_attachments_from_messages(service, query, date_to_read_from, date_to_read_to):
    """
    Gets the attachments from all messages from the first date until the second
    date. if there's no second date so get the messages until today
    (by default of the google API)
    :param service: the gmail API service
    :param date_to_read_from: start date of messages
    :param date_to_read_to: end date of messages
    :return:
    """
    # define the query according to the input
    q = query
    q += " after:" + date_to_read_from
    if date_to_read_to != "":
        q += " before:" + date_to_read_to

    messages = list_messages_matching_query(service, user_id='me', query=q)

    for msg in messages:
        get_attachment_by_msg_id(service, 'me', msg['id'], "Attachments\\")


def downloading_recipet_pdfs(query, date_to_read_from, date_to_read_to=""):
    """
    Connecting to gmail's API and calling the get_attachments function according to
    the given input.
    """
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

    service = build('gmail', 'v1', credentials=creds)

    get_attachments_from_messages(service, query, date_to_read_from, date_to_read_to)
