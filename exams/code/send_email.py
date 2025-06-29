import os.path
import base64
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email import encoders
from email.mime.base import MIMEBase
from mimetypes import guess_type


# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

title="Pedido de Requisição"
path="/Users/danielmdias/docs/_fun_time/00_madrid_prescriber/exams/code/data/output/Output_Test.docx"

def gmail_send_message_with_attachment(title, path):
    """Shows basic usage of the Gmail API.
    Sends an email message with an attachment.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # if not creds or not creds.valid:
    #     if creds and creds.expired and creds.refresh_token:
    #         creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('/Users/danielmdias/docs/_fun_time/00_madrid_prescriber/exams/code/data/required_files/gmail_secrets.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)

        # Create the email content
        message = EmailMessage()
        message.set_content('Para assinar e enviar')

        message['To'] = 'danielmartinhodias@gmail.com'
        message['From'] = 'dmdias.dev@gmail.com'
        message['Subject'] = '{}'.format(title)

        # Attach a file
        file_path = path  # Replace with your file path

        # Guess the file's MIME type
        mime_type, _ = guess_type(file_path)
        mime_type = mime_type or 'application/octet-stream'  # Default to binary

        # Open the file in binary mode
        with open(file_path, 'rb') as f:
            file_data = f.read()

        # Create a MIMEBase instance for the attachment
        attachment = MIMEBase(*mime_type.split('/'))
        attachment.set_payload(file_data)
        encoders.encode_base64(attachment)  # Encode to base64

        # Add the headers for the attachment
        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename="{os.path.basename(file_path)}"',
        )

        # Attach the file to the message
        message.add_attachment(attachment)

        # Encode the entire message as base64
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        # Create the message dictionary for Gmail API
        create_message = {
            'raw': encoded_message
        }

        # Send the email
        send_message = service.users().messages().send(userId='me', body=create_message).execute()
        print(f'Message Id: {send_message["id"]}')

    except HttpError as error:
        print(f'An error occurred: {error}')
        send_message = None

    return send_message


gmail_send_message_with_attachment(title, path)

