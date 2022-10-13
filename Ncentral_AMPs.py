# Ncentral_AMPs.py - parse AMP output emails from Outlook,
# while also keeping track of client_name
# and job details from output email body. Parse data,
# create/update master client_name excel files, move processed email when done.

# Just select success and failure in amp setting;
# do not select to send task output in file

# One small bug that I don't know how to solve yet:
# output from devices that reside in sub-sites of a client show that site
# and only that site as the customer.
# Nowhere in the output does the parent company show.
# So if running an amp on a client with sites,
# be aware of this when looking at the final datafile.

# Author: Josh Smith

import win32com.client
import re
import os
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
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

# REGEX block
# regex to find zip files (Not in use, but keeping in case)
zip_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
# regex to find Client names
cust_regex = re.compile(r'''^.*(Customer: (.*?)) Executed By:''')
# regex to find job type and amp/script name
type_regex = re.compile(r'''^.*Type: (.*?) \[(.*?)\]''')
# regex to find device name
device_regex = re.compile(r'''^.*Device: (.*?) \[''')
# regex to find bde status output
bde_regex = re.compile(r'''(Conversion Status: )(.*?) (Percentage)''')
# regex to find TPM status
tpm_regex = re.compile(r'''oscpresent: (.*?)
                           oscactive: (.*?)
                           oscenabled: (.*?)
                           Result''', re.VERBOSE)


# iterate through all emails and process (Main block)
for msg in list(messages):
    # parse info from email body and organize into variables
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

    # client folder management (not sure if folders are necessary yet)
    ####
    # if os.path.exists(parent_f + client_name) is False:
    #     os.makedirs(parent_f + client_name)
    # client_folder = parent_f + client_name + '\\'
    # logging.debug(f'Client Folder location: {client_folder}')
    ####

    # TODO While keeping track of file's parent company, job,
    #  read output contents,
    #  and update client_name spreadsheet with device and details

    wb = Workbook()
    wb_file = parent_f + f'{client_name}.xlsx'
    # Check if client xlsx exists, if not create, and prep
    if os.path.exists(wb_file) is False:
        wb.save(wb_file)
        wb = load_workbook(wb_file)
        wb_sheet = wb['Sheet']
        wb_sheet.title = 'Encryption'
        headers = [('Device Name', 'TPM Present?', 'TPM Active?',
                   'TPM Enabled?', 'Encryption Status')]
        for i in range(1, 6):
            col = get_column_letter(i)
            wb_sheet.column_dimensions[col].width = 25
        for row in headers:
            wb_sheet.append(row)
        wb.save(wb_file)
    # wb = load_workbook(wb_file)
    # Handle TPM amp and populate variables
    if job_name == 'Windows TPM Monitoring':
        tpm_mo = re.search(tpm_regex, msg.Body)
        if tpm_mo:
            tpm_present = tpm_mo.group(1)
            tpm_active = tpm_mo.group(2)
            tpm_enabled = tpm_mo.group(3)
            logging.debug(f'tpm present?: {tpm_present}')
            logging.debug(f'tpm active?: {tpm_active}')
            logging.debug(f'tpm enabled?: {tpm_enabled}')
    # Handle encryption status check
    elif job_name == 'manage-bde -status':
        bde_mo = re.search(bde_regex, msg.Body)
        if bde_mo:
            encrypt_status = bde_mo.group(2)
            logging.debug(f'Encrypted status: {encrypt_status}')
    # Handle anything else right now
    else:
        print(f'No support added for {job_name} yet. Sorry.')
        logging.debug('End of msg process\n')
        continue

    # TODO for now, support only for TPM checks, BDE/encryption status
    logging.debug('End of msg process\n')
    # msg.Move(outbox)


logging.debug('End of program')