#! python3
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

# TODO for now, support only for TPM checks, BDE/encryption status.
#  Want to add support for on/offboarding status in lieu of
#  relying on asset scans.

# Author: Josh Smith

import win32com.client
import re
import os
import shelve
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %'
                                                '(levelname)s - %'
                                                '(message)s'
                    )

logging.disable(logging.CRITICAL)
logging.debug('Start of program\n')

# Variable initialization
db = 'U:\\Joshua\\Dropbox\\Dropbox\\Python\\Work Stuff\\work_stuff'
with shelve.open(db) as shelf:
    email = shelf['josh_email']
    parent_f = shelf['out_folder']
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders.Item(email).Folders[
    'Inbox'].Folders['Auto Policy']
outbox = outlook.Folders.Item(email).Folders[
    'Inbox'].Folders['Auto Policy'].Folders['Processed']
messages = inbox.Items
err_file = parent_f + 'AMP_errors.txt'
gold = 'FFD966'
red = 'E06666'


# REGEX block
# regex to find zip files (Not in use, but keeping in case)
zip_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
# regex to find Client names
cust_regex = re.compile(r'''^.*(Customer: (.*?))Executed By:''')
# regex to find job type and amp/script name
type_regex = re.compile(r'''^.*Type: (.*?) \[(.*?)\]''')
# regex to find device name
device_regex = re.compile(r'''^.*Device: (.*?)\[''')
# regex to find bde status output
bde_regex = re.compile(r'''(Conversion Status: )(.*?) (Percentage)''')
# regex to find TPM status
tpm_regex = re.compile(r'''oscpresent:(.*?)
                           oscactive:(.*?)
                           oscenabled:(.*?)
                           Result''', re.VERBOSE)


# iterate through all emails and process (Main block)
for msg in list(messages):
    # parse info from email body, organize into variables, and handle errors
    cust_mo = re.search(cust_regex, msg.Body)
    type_mo = re.search(type_regex, msg.Body)
    device_mo = re.search(device_regex, msg.Body)

    if cust_mo:
        client_name = cust_mo.group(2).strip()
        logging.debug(f'Client: {client_name}')
    else:
        with open(err_file, 'a') as file:
            file.write(f'\nNo Customer Detected. '
                       f'Skipping Email Subject: {msg.Subject}...\n')
        continue

    if type_mo:
        job_type = type_mo.group(1).strip()
        logging.debug(f'Job type: {job_type}')
        job_name = type_mo.group(2).strip()
        logging.debug(f'Job name: {job_name}')
    else:
        with open(err_file, 'a') as file:
            file.write(f'\nNo Job Detected. '
                       f'Skipping Email Subject: {msg.Subject}...\n')
        continue

    if device_mo:
        device_name = device_mo.group(1).strip()
        logging.debug(f'Device name: {device_name}')
    else:
        with open(err_file, 'a') as file:
            file.write(f'\nNo Device Detected. '
                       f'Skipping Email Subject: {msg.Subject}...\n')
        continue

    # Added check to handle some occasional false "successes"
    if 'Task did not produce any output.' in msg.Body:
        with open(err_file, 'a') as file:
            file.write(f'\n{client_name}: '
                       f'{device_name} did not produce any output for '
                       f'{job_name}')
        continue

    # Added check for another possible error
    if 'This policy has encountered an exception' in msg.Body:
        with open(err_file, 'a') as file:
            file.write(f'\n{client_name}: '
                       f'{device_name} failed to run '
                       f'{job_name}')
        continue

    # client folder management (not sure if folders are necessary yet)
    ####
    # if os.path.exists(parent_f + client_name) is False:
    #     os.makedirs(parent_f + client_name)
    # client_folder = parent_f + client_name + '\\'
    # logging.debug(f'Client Folder location: {client_folder}')
    ####

    # While keeping track of file's parent company, job, read output contents,
    #  and update client_name spreadsheet with device and details
    wb = Workbook()
    wb_file = parent_f + f'{client_name}.xlsx'

    # Check if client xlsx exists, if not create, and prep
    if os.path.exists(wb_file) is False:
        wb.save(wb_file)
        wb = load_workbook(wb_file)
        wb_sheet = wb['Sheet']
        wb_sheet.title = 'Encryption'
        font_header = Font(size=12, bold=True)
        headers = [('Device Name', 'TPM Present?', 'TPM Active?',
                   'TPM Enabled?', 'Encryption Status')]
        for i in range(1, 6):
            col = get_column_letter(i)
            wb_sheet.column_dimensions[col].width = 25
        for row in headers:
            wb_sheet.append(row)
        for cell in wb_sheet['1:1']:
            cell.font = font_header
        wb_sheet.freeze_panes = 'A2'
        wb.save(wb_file)

    # Open client excel file, get current max row,
    # iterate to check if device already exists

    #TODO put this little detection block in a function so I can pass sheet name
    # and use it for other amps to keep everything in the same workbook

    wb = load_workbook(wb_file)
    sheet = wb['Encryption']
    max_row = sheet.max_row
    device_row = ''
    for i in range(1, max_row + 1):
        cell_data = sheet.cell(row=i, column=1).value
        if device_name in cell_data:
            device_row = i
            break
        else:
            device_row = ''

    # Handle TPM amp and populate spreadsheet
    if job_name == 'Windows TPM Monitoring':
        tpm_mo = re.search(tpm_regex, msg.Body)
        if tpm_mo:
            tpm_present = tpm_mo.group(1).strip()
            tpm_active = tpm_mo.group(2).strip()
            tpm_enabled = tpm_mo.group(3).strip()
            logging.debug(f'tpm present?: {tpm_present}')
            logging.debug(f'tpm active?: {tpm_active}')
            logging.debug(f'tpm enabled?: {tpm_enabled}')
            if device_row == '':
                new_row = [(device_name, tpm_present,
                            tpm_active, tpm_enabled)]
                device_row = sheet.max_row + 1
                for row in new_row:
                    sheet.append(row)
            else:
                sheet.cell(row=device_row, column=2).value = tpm_present
                sheet.cell(row=device_row, column=3).value = tpm_active
                sheet.cell(row=device_row, column=4).value = tpm_enabled

            # highlight any row with no TPM detected
            if tpm_present == 'No TPM Detected':
                for cell in sheet[device_row]:
                    cell.fill = PatternFill(start_color=gold,
                                            end_color=gold,
                                            fill_type='solid')

    # Handle encryption status check
    elif job_name == 'manage-bde -status':
        bde_mo = re.search(bde_regex, msg.Body)
        if bde_mo:
            encrypt_status = bde_mo.group(2).strip()
            logging.debug(f'Encrypted status: {encrypt_status}')
            if device_row == '':
                new_row = [(device_name, '', '', '', encrypt_status)]
                for row in new_row:
                    sheet.append(row)
            else:
                sheet.cell(row=device_row, column=5).value = encrypt_status

    # Handle anything else right now
    else:
        with open(err_file, 'a') as file:
            file.write(f'\nNo support added for {job_name} yet. Sorry. '
                       f'Skipping {client_name}: '
                       f'{device_name}: {job_name}...')
        logging.debug('End of msg process\n')
        continue

    wb.save(wb_file)
    logging.debug('End of msg process\n')
    msg.Move(outbox)

logging.debug('End of program')