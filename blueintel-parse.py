'''
blueintel-parse.py

Uses the O365 module at https://o365.github.io/python-o365/latest/html/index.html

Logs into O365 as the specified user in credentials.ini locates messages from blueintel credparser,
downloads the messages. Then it separates the CSV file attachments from the message, and stores them
in a temporary directory for processing. The CSV attachments are combined together into one file,
then entries are deduplicated, and verified against LDAP for accuracy.

Requires a local ./credentials.ini file, see credentials-example.ini.
'''

from O365 import MSGraphProtocol, Account, FileSystemTokenBackend, Message
from pathlib import Path
from configparser import RawConfigParser
from re import search
import os
import shutil
from fnmatch import fnmatch
import csv
import re
from datetime import datetime
from ldap3 import Server, Connection, ALL, NTLM, ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES, AUTO_BIND_NO_TLS, SUBTREE
from ldap3.core.exceptions import LDAPCursorError
from pprint import pprint

# Functions in order of operations

def pull(account, mailfolder, root):
    print("[+] Pulling credparser emails from acopeland@celgene.com\credparser")

    root = os.path.join('.', root)
    mailbox = account.mailbox()
    mail_folder = mailbox.get_folder(folder_name=mailfolder)
    folder_messages = mail_folder.get_messages(download_attachments=True)

    for message in folder_messages:
        attachments = message.attachments
        for attachment in attachments:
            pattern = '.csv$'
            is_csv = search(pattern, str(attachment))
            if is_csv:
                if not os.path.exists(root):
                    os.makedirs(root)
                attachment.save(location=root)

def parse(parse_file, root, pattern):
    print("[+] Reformating the output and combining the CSV files")

    filenames = []

    # In the attachments folder, find all of the csv files
    for root, dirs, files in os.walk(root):
        for name in files:
            if fnmatch(name, pattern):
                fn = root + "/" + name
                filenames.append(fn)

    # Set the field names for the output file
    fields = ['Source', 'Username', 'Masked Password', 'Destination', 'CSV File']

    # Opening up our new output file
    with open(parse_file, 'w', encoding="UTF-8", newline='') as csvfile:

        # Establish a csvwriter object
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(fields)

        # Open each csv file
        for fn in filenames:
            f = open(fn, encoding="UTF-8")
            f.readline()
            csv_f = csv.reader(f, delimiter=',')

            # Iterate through each row
            for row in csv_f:
                # If the password column is a hash, replace it with "<Password Hash>"
                if re.search('\[HASH\]', str(row[0])) or re.findall(r"(^[a-fA-F\d]{32})", row[2]):
                    row[2] = "<Password Hash>"
                    row.append(fn.split("/")[2])
                    csvwriter.writerow(row)
                elif not row[2] == 'NF' or row[2] == "EMPTY:NONE":
                    row.append(fn.split("/")[2])
                    strlength = len(row[2])
                    masked = int(strlength / 3)
                    slimstr = row[2][masked:strlength - masked]
                    row[2] = "*" * masked + slimstr + "*" * masked
                    csvwriter.writerow(row)

def dedup(parse_file, dedup_file):
    print("[+] Deduplicating entries in the output CSV.")

    lines_seen = set()  # holds lines already seen

    outfile = open(dedup_file, "w", encoding="UTF-8")
    for line in open(parse_file, "r", encoding="UTF-8"):
        line_list = line.split(",")
        if line_list[1] not in lines_seen:  # not a duplicate
            outfile.write(line)
            lines_seen.add(line_list[1])
    outfile.close()

def verify(dedup_file, verify_file, server_name, domain_name, user_name, password):
    print("[+] Verifying entries on the deduplicated CSV.")

    # Setting up LDAP session to Active Directory
    server = Server(server_name, get_info=ALL)
    conn = Connection(server, user='{}\\{}'.format(domain_name, user_name), password=password, authentication=NTLM,
                      auto_bind=True)

    # Establish a csvwriter object
    fields = ['Source', 'Username', 'Masked Passord', 'Destination', 'CSV File', 'Full Name', 'pwdLastSet']
    csvfile = open(verify_file, 'w', encoding="UTF-8", newline='')
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(fields)

    input_file = open(dedup_file, encoding="UTF-8")
    input_file.readline()
    csv_input_file = csv.reader(input_file, delimiter=',')

    # Reading in and iterating through the CSV
    for line in csv_input_file:

        search_account = line[1].split("@")[0]

        # Searching AD for the user in the current CSV file's line
        search_filter = '(&(objectclass=person)(sAMAccountName=' + search_account + ')' + \
                        '(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
        conn.search('DC=celgene,DC=com'.format(domain_name), search_filter=search_filter,
                    attributes=[ALL_ATTRIBUTES, ALL_OPERATIONAL_ATTRIBUTES])

        for user in conn.entries:
            try:
                luser = {'accountName': str(user.sAMAccountName), 'userName': str(user.name),
                         'pwdLastSet': str(user.pwdLastSet)}

                line.append(luser['userName'])
                line.append(luser['pwdLastSet'])

                csvwriter.writerow(line)

            except LDAPCursorError:
                continue

def sendmail(account, verify_file, recipient, sender):

    print("[+] Sending email with attachment.")

    today = datetime.now().strftime('%Y-%m-%d')
    mailbox = account.mailbox()
    m = mailbox.new_message()

    m.to.add(recipient)
    m.subject = 'Credparser Notification: ' + today
    m.sender.address = sender

    body = """
            <html>
            <body>
                <p>Team,</p>
                <p>Attached are the latest notifications for credparser from {today}.
                Please notify the appropriate users.</p>
                <p>---------------------------------</p>
                <p>This message was automatically generated by blueintel-parse.py</p>                
            </body>
            </html>
            """.format(today=today)
    m.body = body
    m.attachments.add(verify_file)
    m.send()

def archive(account, mailfromfolder, mailtofolder):
    print("[+] Moving credparser emails to archive")

    mailbox = account.mailbox()
    mailfromfolder = mailbox.get_folder(folder_name=mailfromfolder)
    folder_messages = mailfromfolder.get_messages(download_attachments=False)
    mailtofolder = mailbox.get_folder(folder_name=mailtofolder)

    for message in folder_messages:
        message.move(mailtofolder)

def cleanup(parse_file, root, dedup_file, verify_file):
    print("[+] Cleaning up temporary files and directories.")

    try:
        os.remove(parse_file)
        os.remove(dedup_file)
        os.remove(verify_file)
        shutil.rmtree(root)
    except OSError as e:
        print("[-] %s - %s." % (e.filename, e.strerror))

if __name__ == '__main__':

    # Reading secrets in from configuration file
    parser = RawConfigParser()
    parser.read('credentials.ini')

    # Getting O365 Token
    client_id = parser.get('credentials', 'client_id')
    client_secret = parser.get('credentials', 'client_secret')

    scopes = ['basic', 'mailbox', 'mailbox_shared', 'users', 'message_all', 'message_all_shared']

    credentials = credentials = (client_id, client_secret)
    token_backend = FileSystemTokenBackend(token_path='.', token_filename='token.txt')
    account = Account(credentials, token_backend=token_backend)

    # Authenticating to O365
    if not account.is_authenticated:
        try:
            account.authenticate(scopes=scopes)
        except Exception as e:
            e.args += (d,)
            raise



    # Getting Active Directory Credentials
    server_name = parser.get('ldap', 'server_name')
    domain_name = parser.get('ldap', 'domain_name')
    user_name = parser.get('ldap', 'user_name')
    password = parser.get('ldap', 'password')

    # Getting mail address info
    pattern = parser.get('meta', 'pattern')
    recipient = parser.get('mail', 'recipient')
    sender = parser.get('mail', 'sender')

    # Setting up file names and directories
    parse_file = 'credparser.csv'
    dedup_file = 'credparser-dedup.csv'
    verify_file = "credparser-" + datetime.now().strftime('%Y-%m-%d') + ".csv"
    mailfromfolder = 'credparser'
    mailtofolder = 'credparserarchive'
    root = "./attachments"

    # Bulk operations
    pull(account, mailfromfolder, root)
    parse(parse_file, root, pattern)
    dedup(parse_file, dedup_file)
    verify(dedup_file, verify_file, server_name, domain_name, user_name, password)
    sendmail(account, verify_file, recipient, sender)
    archive(account, mailfromfolder, mailtofolder)
    cleanup(parse_file, root, dedup_file, verify_file)