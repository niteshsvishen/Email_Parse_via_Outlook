import win32com.client
import sys
import csv  
import re
import os
from datetime import date

def extract(count):
    """Get emails from outlook."""
    items = []
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox
    messages = inbox.Items
    message = messages.GetFirst()

    i = 0
    while message:
        try:
            # tmp["Subject"] = message.Subject
            # tmp["SentOn"] = message.senton.date()
            # tmp["Sender"] = message.SenderEmailAddress
            # tmp["Body"] = message.body 
          # print('---', message.Subject, re.search('red light', message.Subject, re.IGNORECASE), re.search('Red Light', message.Subject, re.IGNORECASE))
          if re.search('red light', message.Subject, re.IGNORECASE):  # or re.search('red light', message.body, re.IGNORECASE)
            items.append([
              message.senton.date(),
              message.SenderEmailAddress,
              message.Subject
            ])
            
        except Exception as ex:
            print("Error processing mail", ex)
        i += 1
        if i < count:
            message = messages.GetNext()
        else:
            return items
    return items


def save_message(items):
  header = ['SentOn', 'Sender', 'Subject']
  os.chmod('outlook_farming_001.csv', 0o777)
  with open('outlook_farming_001.csv', 'w', encoding='UTF8', newline='') as f:
    writer = csv.writer(f)

    # write the header
    writer.writerow(header)

    # write multiple rows
    writer.writerows(items)


def main():
    """Fetch and display top message."""
    items = extract(100)
    save_message(items)


if __name__ == "__main__":
    main()