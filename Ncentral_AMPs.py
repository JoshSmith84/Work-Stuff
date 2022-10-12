# Ncentral_AMPs.py - parse AMP output emails from Outlook,
# download attachments while also keeping track of client and
# job details from output email body.
# Must select to send output file in email when setting the job in ncentral.
# Parse data, create/update master client excel files, delete processed files.
# Author: Josh Smith

import win32com.client
import re
import sys
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %'
                                                '(levelname)s - %'
                                                '(message)s'
                    )

# logging.disable(logging.CRITICAL)
logging.debug('Start of program')

# Variable initialization
parent_f = 'U:\\Joshua\\Work-Stuff\\AMP\\'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders.Item(3).Folders['Inbox'].Folders['Auto Policy']
outbox = outlook.Folders.Item(3).Folders['Inbox'].Folders[
    'Auto Policy'].Folders['Processed']
messages = inbox.Items

# regex to find zip files
zip_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
# regex to find Client names
cust_regex = re.compile(r'''^.*(Customer: (.*?)) -''')
# regex to find job type and amp/script name
type_regex = re.compile(r'''^.*Type: (.*?) \[(.*?)\]''')
# regex to find device name
device_regex = re.compile(r'''^.*Device: (.*?) \[''')


# iterate through all emails in the parent_f
for msg in list(messages):
    # TODO pull info from email body to organize
    cust_mo = re.search(cust_regex, msg.Body)
    type_mo = re.search(type_regex, msg.Body)
    device_mo = re.search(device_regex, msg.Body)
    if cust_mo:
        client = cust_mo.group(2)
        logging.debug(f'Client: {client}')
    if type_mo:
        job_type = type_mo.group(1)
        logging.debug(f'Job type: {job_type}')
        job_name = type_mo.group(2)
        logging.debug(f'Job name: {job_name}')
    if device_mo:
        device_name = device_mo.group(1)
        logging.debug(f'Device name: {device_name}')
    # Iterate through all attachments in each email
    # for atmt in msg.Attachments:
    #     mo = re.search(zip_regex, str(atmt))
    #     if mo:
    #         # If found, save attachment and move email
    #         temp_filename = parent_f + msg.Subject + '_' + atmt.FileName
    #         atmt.SaveAsFile(temp_filename)
    #         print('File Successfully Saved [{}]'.format(temp_filename))
    #         msg.Move(outbox)
    # TODO While keeping track of file's parent company, job, crack open zip.
    # TODO read output contents, and
    #  update client spreadsheet with device and details
    # TODO for now, support only for TPM checks, BDE/encryption status
    logging.debug('End of msg process\n')


logging.debug('End of program')