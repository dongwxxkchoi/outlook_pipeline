import win32com.client
import os
import mimetypes
from email.header import Header

import email.charset
from email.message import EmailMessage

from utils.utils import hwp_to_txt

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

save_folder = os.getcwd()

for i, message in enumerate(inbox.Items):
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
        if attachment.FileName.endswith(".hwp"):
            # convert hwp to txt
            hwp_path = os.path.join(folder_path, attachment.FileName)
            decoded_text = hwp_to_txt(hwp_path)
            txt_path = os.path.join(folder_path, f"{attachment.FileName[:-4]}.txt")
            email_msg.add_attachment(open(txt_path, 'rb').read(), maintype='application', subtype='octet-stream', filename=attachment.FileName)
        elif attachment.FileName.endswith((".pdf", ".docx", ".doc", ".xlsx", ".xls", ".hwp", ".pptx", ".ppt", ".txt", ".zip")):
            # other datatypes
            attachment_path = os.path.join(folder_path, attachment.FileName)
            content_type, encoding = mimetypes.guess_type(attachment_path)
            attachment.SaveAsFile(attachment_path)
            email_msg.add_attachment(open(attachment_path, 'rb').read(), maintype='application', subtype='octet-stream', filename=attachment.FileName)
        else:
            # will not process this attachment files (including jpg)
            continue

    file_name = f"eml/email_{i}.eml"
    file_path = os.path.join(save_folder, file_name)

    with open(file_path, 'wb') as file:
        file.write(email_msg.as_bytes())