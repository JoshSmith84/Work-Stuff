# Ncentral_AMPs.py - parse AMP output emails from Outlook,
# while also keeping track of client_name and job details from output email body.
# Parse data, create/update master client_name excel files, delete processed files.

# Just select success and failure in amp setting;
# do not select to send task output in file

# Author: Josh Smith

import win32com.client
import re
import os
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %'
                                                '(levelname)s - %'
                                                '(message)s'
                    )

# logging.disable(logging.CRITICAL)
logging.debug('Start of program\n')

# Variable initialization
parent_f = 'U:\\Joshua\\Work-Stuff\\AMP\\'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders.Item(3).Folders['Inbox'].Folders['Auto Policy']
outbox = outlook.Folders.Item(3).Folders['Inbox'].Folders[
    'Auto Policy'].Folders['Processed']
messages = inbox.Items

# regex to find zip files (Not in use, but keeping in case)
zip_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
# regex to find Client names
cust_regex = re.compile(r'''^.*(Customer: (.*?)) -''')
# regex to find job type and amp/script name
type_regex = re.compile(r'''^.*Type: (.*?) \[(.*?)\]''')
# regex to find device name
device_regex = re.compile(r'''^.*Device: (.*?) \[''')
# regex to find


# iterate through all emails in the parent_f
for msg in list(messages):
    # pull info from email body to organize
    cust_mo = re.search(cust_regex, msg.Body)
    type_mo = re.search(type_regex, msg.Body)
    device_mo = re.search(device_regex, msg.Body)
    if cust_mo:
        client_name = cust_mo.group(2)
        logging.debug(f'Client: {client_name}')
    else:
        print(f'No Customer Detected. Skipping Email Subject: {msg.Subject}...')
        continue
    if type_mo:
        job_type = type_mo.group(1)
        logging.debug(f'Job type: {job_type}')
        job_name = type_mo.group(2)
        logging.debug(f'Job name: {job_name}')
    else:
        print(f'No Job Detected. Skipping Email Subject: {msg.Subject}...')
        continue
    if device_mo:
        device_name = device_mo.group(1)
        logging.debug(f'Device name: {device_name}')
    else:
        print(f'No Device Detected. Skipping Email Subject: {msg.Subject}...')
        continue

    # With info, keep output files organized. Adding check/create now for
    # future support options

    # client folder management
    if os.path.exists(parent_f + client_name) is False:
        os.makedirs(parent_f + client_name)
    client_folder = parent_f + client_name + '\\'
    logging.debug(f'Client Folder location: {client_folder}')

    # Iterate through all email
    # for atmt in msg.Attachments:
    #
    #         #msg.Move(outbox)
    # TODO While keeping track of file's parent company, job,
    #  read output contents,
    #  and update client_name spreadsheet with device and details
    # TODO for now, support only for TPM checks, BDE/encryption status
    logging.debug('End of msg process\n')


logging.debug('End of program')