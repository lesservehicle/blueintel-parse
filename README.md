# blueintel-parse.py

Uses the O365 module at https://o365.github.io/python-o365/latest/html/index.html

Logs into O365 as the specified user in `credentials.ini`, locates messages from blueintel credparser (https://github.com/trbpnd/bimeta) in the specified folder, and downloads the messages. 

Then it separates the CSV file attachments from the message, and stores them in a temporary directory for processing. 

The CSV attachments are then combined together into one file, entries are deduplicated, and verified against LDAP for accuracy. An email to a sender specified in the ini file is sent with the merged and verified file as an attachment, and all of the generated temporary files are deleted.

Requires a local ./credentials.ini file, see the included credentials-example.ini.