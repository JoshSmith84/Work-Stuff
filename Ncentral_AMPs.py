#     Function to download all attachments in a folder in Outlook (in_ofolder).
#     The item index below may be unique to outlook. One must refactor this
#     and figure out the index of their mailbox in Outlook
#     especially if there are multiple accounts.
#
#     Once found, the attachment is downloaded to the out_folder param,
#     and the email is moved from in_ofolder param to the out_ofolder param

import win32com.client
import re



# Variable initialization
outlook = win32com.client.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders.Item(3).Folders['Inbox'].Folders[in_ofolder]
outbox = outlook.Folders.Item(3).Folders[
    'Inbox'].Folders[out_ofolder]
messages = inbox.Items
# regex looking for .zip files
amp_out_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
# iterate through all emails in the folder
for msg in list(messages):
    # TODO pull info from email body to organize
    # Start with Customer regex
    print(msg.Body)
    cust_regex = re.compile(r'''^.*(Customer: (.*?)) -''')
    type_regex =
    cust_mo = re.search(cust_regex, msg.Body)
    if cust_mo:
        client = cust_mo.group(2)
    # Iterate through all attachments in each email
    # for atmt in msg.Attachments:
    #     mo = re.search(amp_out_regex, str(atmt))
    #     if mo:
    #         # If found, save attachment and move email
    #         temp_filename = out_folder + msg.Subject + '_' + atmt.FileName
    #         atmt.SaveAsFile(temp_filename)
    #         print('File Successfully Saved [{}]'.format(temp_filename))
    #         msg.Move(outbox)


folder = 'U:\\Joshua\\Work-Stuff\\AMP\\'
