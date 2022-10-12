# credit to Hridai for base function. I made a few edits.
# Such as leveraging regex
# https://github.com/Hridai/Automating_Outlook/blob/master/ol_script.py

import win32com.client
import re

def amp_output_pull(outdest, olreadfolder, olprocessedfolder):
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders.Item(3).Folders['Inbox'].Folders[olreadfolder]
    outbox = outlook.Folders.Item(3).Folders[
        'Inbox'].Folders[olprocessedfolder]
    messages = inbox.Items
    # changed this to use regex to detect file type.
    amp_out_regex = re.compile(r"""^(.*?)(\.)(zip)$""")
    for msg in list(messages):
        for atmt in msg.Attachments:
            print(str(atmt))
            mo = re.search(amp_out_regex, str(atmt))
            if mo:
                temp_filename = outdest + msg.Subject + '_' + atmt.FileName
                atmt.SaveAsFile(temp_filename)
                print('File Successfully Saved [{}]'.format(temp_filename))
                msg.Move(outbox)


folder = 'C:\\temp\\test\\'

amp_output_pull(folder, 'Auto Policy', 'Processed')