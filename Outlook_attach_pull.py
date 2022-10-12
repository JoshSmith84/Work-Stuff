# credit to Hridai for base function. I made a few edits.
# Such as leveraging regex
# https://github.com/Hridai/Automating_Outlook/blob/master/ol_script.py

import win32com.client
import re


def amp_output_pull(out_folder: str, in_ofolder: str, out_ofolder: str) -> None:
    """
    Function to download all attachments in a parent_f in Outlook (in_ofolder).
    The item index below may be unique to outlook. One must refactor this
    and figure out the index of their mailbox in Outlook
    especially if there are multiple accounts.
    Once found, the attachment is downloaded to the out_folder param,
    and the email is moved from in_ofolder param to the out_ofolder param
    Needed imports to use this: win32.client, re
    :param out_folder: Folder on local PC to save attachments to
    :param in_ofolder: Outlook parent_f to search incoming emails.
    :param out_ofolder: Outlook parent_f to send email after it is processed
    """
    # Variable initialization
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders.Item(3).Folders['Inbox'].Folders[in_ofolder]
    outbox = outlook.Folders.Item(3).Folders[
        'Inbox'].Folders[out_ofolder]
    messages = inbox.Items
    # regex looking for .zip files
    amp_out_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
    # iterate through all emails in the parent_f
    for msg in list(messages):
        # Iterate through all attachments in each email
        for atmt in msg.Attachments:
            mo = re.search(amp_out_regex, str(atmt))
            if mo:
                # If found, save attachment and move email
                temp_filename = out_folder + msg.Subject + '_' + atmt.FileName
                atmt.SaveAsFile(temp_filename)
                print('File Successfully Saved [{}]'.format(temp_filename))
                msg.Move(outbox)