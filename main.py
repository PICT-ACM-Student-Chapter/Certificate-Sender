import base64
import csv
import mimetypes
import os.path
import pickle
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from PIL import Image, ImageDraw, ImageFont
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import config

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://mail.google.com/']


def main():
    """
    Validates credentials and initialise the service
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

    return build('gmail', 'v1', credentials=creds)


def create_message_with_attachment(
        sender, to, subject, message_text, file):
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
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'image':
        fp = open(file, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(file, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(file)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}


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
        message = (service.users().messages().send(userId=user_id, body=message)
                   .execute())
        print('Message Id: ' + str(message['id']))
        return message
    except Exception as err:
        print('An error occurred: ' + str(err))


def generate_certificate(name, team_name, file_name):
    """
    Generates a certificate from template.
    """

    # update this as needed
    width = 3508
    image = Image.open(config.certificate_template)
    header_font = ImageFont.truetype('TIMES.ttf', 140)
    body_font = ImageFont.truetype('TIMES.ttf', 70)
    draw = ImageDraw.Draw(image)
    w, h = draw.textsize(name.upper(), font=header_font)
    draw.text(((width - w) / 2, 1260), text=name.upper(), fill="black", font=header_font)
    draw.text(xy=(1750, 1470), text=name, fill=(0, 0, 0), font=body_font)
    draw.text(xy=(1100, 1550), text=team_name, fill=(0, 0, 0), font=body_font)

    image.save(file_name)
    print('Certificate generated for =>' + name)


if __name__ == '__main__':
    service_obj = main()
    data = config.data_file
    with open(data, 'r') as csv_file:
        # creating a csv reader object
        csv_reader = csv.reader(csv_file)

        # extracting each data row one by one
        for row in csv_reader:
            first_name = row[0]
            last_name = row[1]
            email = row[2]
            team_name = row[3]
            file_name = 'certificates/' + team_name + '_' + first_name + '_' + last_name + '.jpg'
            generate_certificate(first_name + ' ' + last_name, team_name, file_name)

            body = config.body

            message_obj = create_message_with_attachment(
                'me',
                email,
                config.subject,
                body,
                file_name
            )
            send_message(service_obj, 'me', message_obj)
