import win32com.client
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import Header

import email.charset
from email.message import EmailMessage

import extract_msg

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

save_folder = os.getcwd()

for i, message in enumerate(inbox.Items):
    # email_msg = MIMEMultipart('alternative')
    email_msg = EmailMessage()

    email.charset.add_charset('utf-8', email.charset.SHORTEST, None, 'utf-8')
    email_msg.set_charset('utf-8')

    email_msg['Subject'] = message.Subject
    email_msg['From'] = message.SenderEmailAddress
    email_msg['To'] = message.To
    email_msg['Date'] = message.ReceivedTime.strftime("%a, %d %b %Y %H:%M:%S %z")

    email_msg.set_content(message.Body, charset='utf-8')
    
    folder_path = 'C:/Users/dwcho/langchain/msg/attachments'

    for attachment in message.Attachments:
        attachment_path = os.path.join(folder_path, attachment.FileName)
        attachment.SaveAsFile(attachment_path)
        email_msg.add_attachment(open(attachment_path, 'rb').read(), maintype='application', subtype='octet-stream', filename=attachment.FileName)

    file_name = f"eml/email_{i}.eml"
    file_path = os.path.join(save_folder, file_name)

    with open(file_path, 'wb') as file:
        file.write(email_msg.as_bytes())

    
#     subject = message.Subject

#     file_name = subject[:10]
#     # file_name = "".join(char for char in subject if char.isalnum())
#     file_path = os.path.join(save_folder,f"data/{file_name}_{i}.msg")

#     message.SaveAs(file_path, 3)

# msgs = os.listdir('data/')
# os.chdir('data')
# for msg in msgs:
#     msg_data = extract_msg.Message(msg)
#     print(msg_data.subject)
