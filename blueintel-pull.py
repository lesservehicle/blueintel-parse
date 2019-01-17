'''
blueintel-pull.py

Uses the O365 module at https://o365.github.io/python-o365/latest/html/index.html

Logs into O365 as the specified user in credentials.ini
locates messages from blueintel credparser, downloads the messages.
Then it separates the CSV file attachments from the message, and
stores them in a temporary (date and timestamp) directory for processing.
'''


from O365 import MSGraphProtocol, Connection, Folder, MailBox, Message, MessageAttachment, Account
from pathlib import Path
import re
from configparser import RawConfigParser

parser = RawConfigParser()
parser.read('credentials.ini')

client_id = parser.get('credentials', 'client_id')
client_secret = parser.get('credentials', 'client_secret')
token_file = Path('./o365_token.txt')

my_protocol = MSGraphProtocol()
my_credentials = (client_id, client_secret)
my_scopes=['basic', 'mailbox', 'mailbox_shared', 'users', 'message_all', 'message_all_shared']

my_account = Account(credentials=my_credentials, protocol=my_protocol')

if not my_account.connection.check_token_file():
    my_account.authenticate(scopes=my_scopes)
    #my_account.connection.get_authorization_url(requested_scopes=my_scopes)
else:
    my_account.connection.get_session(token_path=token_file)

mailbox = my_account.mailbox()
mail_folder = mailbox.get_folder(folder_name='credparser')
my_account.con.refresh_token()

folder_messages = mail_folder.get_messages(download_attachments=True)

for message in folder_messages:
    attachments = message.attachments
    for attachment in attachments:
        pattern = '.csv$'
        is_csv = re.search(pattern,str(attachment))
        if is_csv:
            attachment.save(location='.\\attachments\\')