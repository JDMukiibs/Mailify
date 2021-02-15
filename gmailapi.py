import pickle
import os.path
import base64
from apiclient import errors
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# Function to create the Service
def create_service(secret_file, api_name, api_version, *scopes):
    """Creates a Gmail api service instance that will be used to connect and get authorization from the Gmail-API"""
    SCOPES = [scope for scope in scopes[0]]  # Set the scope to the specific use case which is to send mail.

    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, we are logged in again
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(secret_file, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build(api_name, api_version, credentials=creds)
    return service


def create_message(sender, to, subject, message_text, file):
    """Create a message for an email.

      Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.
        file: The path to the file to be attached.

      Returns:
        An object containing a base64url encoded email object.
      """
    message = MIMEMultipart()
    message['To'] = to
    message['From'] = sender
    message['Subject'] = subject

    # Add body to the email
    message.attach(MIMEText(message_text))

    # I made mine specific because my CV is in pdf format.
    # If you want to use the more general and more encompassing style for other file types,
    # check the code by Google at: https://developers.google.com/gmail/api/guides/sending
    # MIMEApplication is the right choice when working with pdfs which was my case because it helps to read it in octet-stream
    # Open the pdf file and read the binary data in rb mode.
    msg = MIMEApplication(open(file, 'rb').read())
    # Add a header to the attachment
    msg.add_header('Content-Disposition', 'attachment', filename=file)

    # Add attachment to MIME message object and then convert it to string
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def send_message(service, user_id, message):
    """Send an email message.

      Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message: Message to be sent.

      Returns:
        Sent Message.
      """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print("Message Id: {}".format(message['id']))
        return message
    except errors.HttpError as error:
        print("An error occurred: {}".format(error))
