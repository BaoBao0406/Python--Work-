#! python3
# RevenueEmailAttachment.py - Search for specific email subject in Outlook, move the email 
# and download the attachment to specific location

from win32com.client import Dispatch
import datetime, os.path, re

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")
msgs = inbox.Items

# Path for the Attachment to be saved
AttachPath = 'I:\\10-Sales\\Personal Folder\\Admin & Assistant Team\\Patrick Leong\\Python Code\\RevenueEmailAttachement\\'

# Date Range from last three days
d = (datetime.date.today() - datetime.timedelta (days=3)).strftime("%d-%m-%y")


# Outlook folder name to be moved to 
MailFolder = 'Work Related'
# Search for Personal Folder and identify the loction
root_folder = outlook.Folders.Item(3)
for folder in root_folder.Folders:
    if folder.Name == MailFolder:
        donebox = folder

# Search in inbox for last three days
msgs = msgs.Restrict("[ReceivedTime] >= '" + d +"'")

# Function for attachment if email subject contain keyword
def DownloadAttach():
    for att in msg.Attachments:
        path = AttachPath + att.Filename
        att.SaveAsFile(path)
        # TODO: Search for file with the same filename 

# Loop for all email within the Date Range with the keyword
for msg in msgs:
    # Search for keywords that contain
    msgKeyWord = re.compile(r'Testing Python')
    msgSearch = msgKeyWord.search(msg.Subject)
    print(msg.Subject)
    if (msgSearch != None) is True:
        # Run Function
        DownloadAttach()
        # Move the email to specific folder
        msg.Move(donebox)
