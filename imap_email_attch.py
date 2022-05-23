import imaplib
import email
from email.parser import HeaderParser
import os
import base64
from datetime import datetime
from getpass import getpass
#import environ
#env = environ.Env()
#environ.Env.read_env()

server = 'outlook.office365.com' #mail server
#user = env('EMAIL_USER')    #EMAIL ID
user =input("Enter Email ID :")
#password = env('EMAIL_PASSWORD') #EMAIL PASSWORD
password = getpass("Enter Email Password :")
#outputdir = env('OUTPUT_DIR')
#outputdir = 'D:\python\test_email_python' #Windows dir
outputdir = '/mnt/d/python/test_email_python' #Linux dir
keywords_list = ["testai"]

def connect(server, user, password):
    imap_conn = imaplib.IMAP4_SSL(server)
    imap_conn.login(user=user, password=password)
    return imap_conn

def delete_email(instance, email_id):
    typ, delete_response = instance.fetch(email_id, '(FLAGS)')
    typ, response = instance.store(email_id, '+FLAGS', r'(\Deleted)')
    print(delete_response)
    print(response)

#def download_attachmet (instance, email_id)
    #code to save attachment in local folder

def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)

#'filename_2020_08_12-03:29:22_AM'

    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            date = datetime.now().strftime("%Y_%m_%d-%I-%M-%S_%p")
            pre_full_filename = part.get_filename()
            filename, file_extension = os.path.splitext(pre_full_filename)
            final_full_filename = filename+date+file_extension #Split the files name/extenstion, inserted the date and then combined all three
    #    print(f"filename_{date}")
            open(outputdir + '/' + final_full_filename, 'wb').write(part.get_payload(decode=True))

""" Try for download"""
for  keyword in keywords_list:
    conn = connect(server, user, password)
    conn.select('Test1')
    typ, msg = conn.search(None, '(SUBJECT "' + keyword + '")')
    print(msg)
    msg = msg[0].split()
    for email_id in msg:
        print(email_id)
    downloaAttachmentsInEmail (conn, email_id, outputdir)
print('Downloaded')