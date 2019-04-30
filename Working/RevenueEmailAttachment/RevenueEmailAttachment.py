#! python3
# RevenueEmailAttachment.py - Search for specific email subject in Outlook, move the email 
# and download the attachment to specific location

from win32com.client import Dispatch
import datetime, os.path, re, hashlib
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")
msgs = inbox.Items


# Path for the Attachment to be saved
AttachPath = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\RevenueEmailAttachement\\'
# Outlook folder name to be moved to 
MailFolder = 'Work Related'
# Keyword to search for email subject
msgKeyWord = re.compile(r'testing python')

# Date Range from last three days
d = (datetime.date.today() - datetime.timedelta (days=5)).strftime("%d-%m-%y")

# Search for Personal Folder and identify the loction
root_folder = outlook.Folders.Item(3)
for folder in root_folder.Folders:
    if folder.Name == MailFolder:
        donebox = folder

# Search in inbox for last three days
msgs = msgs.Restrict("[ReceivedTime] >= '" + d +"'")

# Use MD5 to search for two files if contain the same data
def MD5(file):
    hasher = hashlib.md5()
    with open(file, 'rb') as afile:
        buf = afile.read()
        hasher.update(buf)
    return hasher.hexdigest()

# Function to download attachment if email subject contain keyword
def DownloadAttach():
    for att in msg.Attachments:
        FileCopied = False
        # Search for file with the same filename 
        for file in os.listdir(AttachPath):
            if att.Filename == file:
                temp = os.getcwd() + '\\(temp)' + att.Filename
                att.SaveAsFile(temp)
                new_md5 = MD5(temp)
                old_md5 = MD5(file)
                if new_md5 != old_md5:
                    # Rename the filename if the file content is not the same
                    path = AttachPath + '(Updated)' + att.Filename
                    att.SaveAsFile(path)
                FileCopied = True
                os.unlink(temp)
        # Copy file to path if current location does not have the same file
        if FileCopied == False:
            path = AttachPath + att.Filename
            att.SaveAsFile(path)

MsgToMove = []
# Loop for all email within the Date Range with the keyword
for msg in msgs:
    # Search for keywords in email subject
    msgSearch = msgKeyWord.search((msg.Subject).lower())
    if (msgSearch == None) is False:
        DownloadAttach()
        MsgToMove.append(msg)
        msgSearch = 'None'
# Move email to specific folder
for msg in MsgToMove:
    msg.Move(donebox)
