import win32com.client
import csv
import pathlib
import os
from datetime import datetime, timedelta

#Obtain all the emails and create list of dictionaries
rows = [] #All emails
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
print("Sent messages")
# Folder reference https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
for folder in mapi.GetDefaultFolder(6).Folders:
    for message in folder.Items:
        folderName = folder.Name
        sender = message.SenderEmailAddress
        sender_domain = sender.split('@')[1]
        subject = message.Subject
       # print("Folder: " + folderName + " Sender: " + sender + " Subject: " + subject)
        rows.append({'Folder': folderName, 'SenderDomain': sender_domain, 'Sender': sender,  'Subject':subject })

#Create the csv file
csv_path = pathlib.Path(os.path.dirname(os.path.realpath(__file__))) / "emails.csv" #Get current path and create new file
fieldnames = ['Folder', 'SenderDomain', 'Sender', 'Subject']
with open(csv_path, 'w', encoding='UTF8', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)
