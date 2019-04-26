#blueintel-parse.py

Uses the O365 module at https://o365.github.io/python-o365/latest/html/index.html

Logs into O365 as the specified user in `credentials.ini`, locates messages from blueintel credparser in the specified folder, and downloads the messages. 

Then it separates the CSV file attachments from the message, and stores them in a temporary directory for processing. 

The CSV attachments are then combined together into one file, entries are deduplicated, and verified against LDAP for accuracy.

Requires a local ./credentials.ini file, structured as follows:

```
[credentials]
client_id = <id>
client_secret = <secret>

[ldap]
server_name = <your LDAP server>
domain_name = <your domain>
user_name = <your username>
password = <your password>
```

##TO DO:
- Update the O365 portion of the script to stop using the deprecated 'token_path' parameter.
- Add a function to include a mail merge.
