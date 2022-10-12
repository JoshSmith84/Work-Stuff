# credit to Hridai for base function. I made a few edits.
# Such as leveraging regex
# https://github.com/Hridai/Automating_Outlook/blob/master/ol_script.py

import win32com.client

def run_ol_Script(outdest, filefmt, olreadfolder, olprocessedfolder):
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders.Item(1).Folders['Inbox'].Folders[olreadfolder]
    outlook.Folders.Item(1).Folders[
        'Inbox'].Folders[olreadfolder].Folders[olprocessedfolder]
    messages = inbox.Items
    #TODO change this to use regex to detect file type.
    for msg in list(messages):
        for atmt in msg.Attachments:
            if filefmt == 'blank' or str.lower(_right(atmt.FileName, len(filefmt))) == str.lower(filefmt):
                temp_fileName = outdest + msg.Subject + '_' + atmt.FileName
                atmt.SaveAsFile(temp_fileName)
                print('File Successfully Saved [{}]'.format(temp_fileName))
                msg.Move(outlook.Folders.Item(1).Folders['Inbox'].Folders[olreadfolder].Folders[olprocessedfolder])

